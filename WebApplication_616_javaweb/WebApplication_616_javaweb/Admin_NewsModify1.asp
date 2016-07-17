<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改文章</title>
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
}
</script>
</head>
<body leftmargin="5" topmargin="10">
<%
dim Rs,sqlStr,id
id=request.querystring("newsid")
set Rs=server.createobject("adodb.recordset")
sqlStr="select * from QDnews where newsid=" & clng(id)
Rs.open sqlStr,conn,1,1
if not Rs.eof then
%>
<form method="POST" name="myform" onSubmit="return CheckForm();" action="Admin_NewsSave1.asp?action=Modify">
  <table border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr class="title">
      <td height="22" align="center"><b>修 改 文 章</b></td>
    </tr>
    <tr align="center">
      <td>
	<table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属栏目：</strong></td>
            <td>
             <table width="100%" border="0" cellspacing="0" cellpadding="0">
               <tr>
                 <td><select name='ClassID' size=1><%
                 dim CRs,Csql
                 set CRs=server.createobject("adodb.recordset")
                 Csql="select * from QDclass"
                 CRs.open Csql,conn,1,1
                 do while not CRs.eof
                 	if CRs("ClassID")=Rs("ClassID") then
                 		response.write "<option value="&CRs("ClassID")&" selected>"&CRs("ClassName")&"</option>"
                 	else
	                 	response.write "<option value="&CRs("ClassID")&">"&CRs("ClassName")&"</option>"
                 	end if
                 	CRs.movenext
                 loop
                 CRs.close
				 %></select>                 </td>
               </tr>
             </table>			</td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章标题：</strong></td>
            <td><input name="Title" type="text"
           id="Title" value="<%=Rs("Title")%>" size="50" maxlength="255">
              <font color="#FF0000">*</font>            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>来源：</strong></td>
            <td> 
              <input name="source" type="text"
            value="<%=Rs("source")%>" size="20" maxlength="30"> 
&nbsp;&nbsp;&nbsp;            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>作者：</strong></td>
            <td> 
              <input name="author" type="text"
           value="<%=Rs("author")%>" size="20" maxlength="30"> 
&nbsp;&nbsp;&nbsp;            </td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>文章内容：</strong></p>
              <p align="left"><font color="#006600">                &middot;　换行请按Shift+Enter</font><br>
              <font color="#006600">&middot;　另起一段请按Enter</font></p></td>
            <td><textarea name="Content" style="display:none"><%=Rs("Content")%></textarea> 
              <iframe ID="editor" src="WebEditor/ewebeditor.htm?id=content&style=coolblue&savepathfilename=myText3" frameborder=1 scrolling=no width="600" height="405"></iframe> </td>
          </tr>
          <tr class="tdbg">
            <td align="right" valign="middle"><strong>点击次数：</strong></td>
            <td colspan="2"><input name="hits" type="text" id="<%=rs("hits")%>" value="1" size="5"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <div align="center"> 
    <p> 
      <input name="newsID" type="hidden"  value="<%=Rs("newsID")%>">
      <input
  name="Save" type="submit"  id="Save" value="保存修改结果" style="cursor:hand;">
      &nbsp; 
      <input name="Cancel" type="reset" id="Cancel" value=" 取 消 " style="cursor:hand;">
    </p>
  </div>
</form>
</body>
</html>
<%
end if
Rs.close
call CloseConn()
%>