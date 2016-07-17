<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>添加文章（简洁模式）</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<script language = "JavaScript">
function AddItem(strFileName){
  //document.myform.IncludePic.checked=true;
  //document.myform.DefaultPicUrl.value=strFileName;
}
function CheckForm()
{
  if (document.myform.Title.value=="")
  {
    alert("文章标题不能为空！");
	document.myform.Title.focus();
	return false;
  }
}
function doChange(objText, objDrop){
	if (!objDrop) return;
	var str = objText.value;
	var arr = str.split("|");
	objDrop.length=0;
	for (var i=0; i<arr.length; i++){
		objDrop.options[i] = new Option(arr[i], arr[i]);
	}
	objDrop.options[i] = new Option('不显示图片','');
}
</script>
</head>

<body>
<form method="POST" name="myform" onSubmit="return CheckForm();" action="Admin_NewsSave1.asp" target="_self">
  <table border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr class="title">
      <td height="22" align="center"><b>添 加 新 闻</b></td>
    </tr>
    <tr align="center">
      <td>
	<table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属栏目：</strong></td>
            <td>
            <%
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
            %>            </td>
            <td></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章标题：</strong></td>
            <td colspan="2"><input name="Title" type="text" size="50" maxlength="255"> 
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>来源：</strong></td>
            <td colspan="2"> <input name="source" type="text" value="本站原创"
           size="50" maxlength="100">            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>作者：</strong></td>
            <td colspan="2"> <input name="Author" type="text" value="金豫科技"
           size="50" maxlength="100">            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>文章内容：</strong></p>
              <p align="left"><font color="#006600">                &middot;　换行请按Shift+Enter</font><br>
                <font color="#006600">&middot;　另起一段请按Enter</font><br>
              </p></td>
            <td colspan="2"><textarea name="Content" style="display:none"></textarea> 
              <iframe ID="editor" src="WebEditor/ewebeditor.htm?id=content&style=coolblue&savepathfilename=myText3" frameborder=1 scrolling=no width="600" height="405"></iframe> </td>
          </tr>
          <tr class="tdbg">
            <td align="right" valign="middle"><strong>点击次数：</strong></td>
            <td colspan="2"><input name="hits" type="text" id="hits" value="1" size="5"></td>
          </tr>
          <!--<tr class="tdbg">
            <td align="right" valign="middle"><strong>首页图片：</strong></td>
            <td colspan="2">
			<input id=myText3 style="width:200px" onChange="doChange(this,defaultpic)" type="hidden">&nbsp;<select id=defaultpic name="defaultpic" size=1 style="width:350px"></select>
			</td>
          </tr>-->
        </table>
      </td>
    </tr>
  </table>
  <div align="center"> 
    <p>
      <input name="Action" type="hidden" id="Action" value="Add">
      <input
  name="Add" type="submit"  id="Add" value=" 添 加 "style="cursor:hand;">
      &nbsp; 
      &nbsp; 
      <input name="Cancel" type="reset" id="Cancel" value=" 取 消 " style="cursor:hand;">
    </p>
  </div>
</form>
</body>
</html>
<%
call CloseConn()
%>