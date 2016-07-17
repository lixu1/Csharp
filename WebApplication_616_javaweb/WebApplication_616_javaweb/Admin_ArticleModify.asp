<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%

dim Rs,sqlStr,Rs1
dim sql
dim count
dim ArticleID,FoundErr
FoundErr=False
ArticleID=trim(request.QueryString("ArticleID"))
if ArticleID="" then 
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>请指定要修改的产品ID</li>"
	call WriteErrMsg()
	call CloseConn()
	response.end
else
	ArticleID=Clng(ArticleID)
end if
set Rsm=server.CreateObject("adodb.recordset")
sqlStr="select * from Article where ArticleID=" & ArticleID
Rsm.open sqlStr,conn,1,1
if FoundErr=True then
	call WriteErrMsg()
else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加产品</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<script language = "JavaScript">
function AddItem(strFileName){
  //document.myform.IncludePic.checked=true;
  document.myform.DefaultPic.value=strFileName;
}
function CheckForm()
{
  if (document.myform.Title.value=="")
  {
    alert("产品名称不能为空！");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("产品介绍不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
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

	var arr=new Array()
	<%
		set rs=server.CreateObject("adodb.recordset")
		sql="select * from SmallClass order by SmallID"
		rs.open sql,conn,1,1
		if not rs.eof then
			i=0
			do until rs.eof
			bigName=conn.execute("select BigName from BigClass where BigID="&rs("BigID"))(0)
	%>
	arr[<%=i%>]=["<%=rs("SmallName")%>","<%=BigName%>","<%=rs("SmallID")%>"]<%
				i=i+1
				rs.movenext
			loop
		end if
		rs.close
	%>
	function changeSub(bigName){
		var j=0;
		var subClass = document.getElementById("SmallClass");
		subClass.options.length=0;
		for(i=0;i<arr.length;i++){
			if(arr[i][1]==bigName){
				subClass.options.add(new Option(arr[i][0],arr[i][2]));
				j++;
			}
		}
		if(subClass.options.length==0){
			subClass.options.add(new Option("无小分类",""));
			subClass.style.display="none";
		}else{
			subClass.style.display="";
		}
	}
//-->
</script>
</head>
<body>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From article where articleid="&articleID, conn,3,3
%>
<form method="POST" name="myform" onSubmit="return CheckForm();" action="Admin_ArticleSave.asp" target="_self">
  <table border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr class="title">
      <td height="22" align="center"><b>修 改 图 文 信 息</b></td>
    </tr>
    <tr align="center">
      <td>
	<table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr class="tdbg">
            <td align="right"><strong>类别：</strong></td>
            <td colspan="2"><select name="BigClass" id="BigClass" onChange="changeSub(this.options[this.selectedIndex].text)">
                      <%
		set rs1=server.CreateObject("adodb.recordset")
		sql="select * from BigClass"
		rs1.open sql,conn,1,1
		if not rs1.eof then
			do until rs1.eof
	%>
                      <option value="<%=rs1("BigID")%>"<%if rs("bigID")=rs1("bigID") then response.Write(" selected")%>><%=rs1("bigName")%></option>
                      <%
				rs1.movenext
			loop
		end if
		rs1.close
		set rs1=nothing
	%>
                    </select>
					<select name="SmallClass" id="SmallClass">
					<%if rs("SmallID")<>0 then%>
						<option value="<%=rs("SmallID")%>"><%=getSmallName(rs("SmallID"))%></option>
					<%else%>
                        <option>无小类别</option>
					<%end if%>
                    </select></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>名称：</strong></td>
            <td colspan="2"><input name="Title" type="text" id="Title" value="<%=Rsm("title")%>" size="50" maxlength="255"> 
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>介绍：</strong></p>
              <p align="left"><font color="#006600">                &middot;　换行请按Shift+Enter</font><br>
              <font color="#006600">&middot;　另起一段请按Enter</font></p></td>
            <td colspan="2"><textarea name="Content" style="display:none"><%=Rsm("content")%></textarea> 
              <iframe ID="editor" src="WebEditor/ewebeditor.htm?id=content&style=coolblue&savepathfilename=myText3" frameborder=1 scrolling=no width="600" height="405"></iframe> </td>
          </tr>
          <tr class="tdbg">
            <td align="right" valign="middle"><strong>缩略图：</strong></td>
            <td width="200"><input name="DefaultPic" type="text" id="DefaultPic" value="<%=Rsm("defaultpic")%>" style="width:200px" /></td>
            <td width="395"><iframe src="Upload_Article1.asp" frameborder="0" style="height:25px; width:260px;"></iframe></td>
          </tr>
          <tr class="tdbg">
            <td align="right" valign="middle"><strong>点击次数：</strong></td>
            <td colspan="2"><input name="hits" type="text" id="hits" value="1" size="5"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <div align="center"> 
    <p> 
      <input name="Action" type="hidden" value="Modify"><input type="hidden" value="<%=Rsm("ArticleID")%>" name="ArticleID">
      <input
  name="Add" type="submit"  id="Add" value=" 修 改 " onClick="document.myform.action='Admin_ArticleSave.asp';document.myform.target='_self';" style="cursor:hand;">
      &nbsp; 
      &nbsp;
      <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_ArticleManage.asp'" style="cursor:hand;">
    </p>
  </div>
</form>
</body>
</html>
<%
end if
Rsm.close

function getSmallName(id)
	set ors=server.CreateObject("adodb.recordset")
	sql="select * from SmallClass where SmallID="&id
	ors.open sql,conn,1,1
	getSmallName=ors("SmallName")
	ors.close
	set ors=nothing
end function

call CloseConn()
%>