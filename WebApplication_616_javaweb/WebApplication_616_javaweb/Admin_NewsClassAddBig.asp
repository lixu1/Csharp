<%@language=vbscript codepage=936 %>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<%
dim Action,BigClassName,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章大类名不能为空！</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From NewsClass Where ClassName='" & BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>文章大类“" & BigClassName & "”已经存在！</li>"
		else
    	 	rs.addnew
     		rs("ClassName")=BigClassName
    	 	rs.update
     		rs.Close
	     	set rs=Nothing
    	 	call CloseConn()
			Response.Redirect "Admin_NewsClassManage.asp"  
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
%>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("大类名称不能为空！");
    document.form1.BigClassName.focus();
    return false;
  }
}
</script>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><table width="80%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="25" align="center" valign="top"> <br>
            <strong>新 闻 类 别 设 置</strong><br> 
            <form name="form1" method="post" action="Admin_NewsClassAddBig.asp" onsubmit="return checkBig()">
              <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
                <tr bgcolor="#A4B6D7" class="title"> 
                  <td height="20" colspan="2" align="center"><strong>添加大类</strong></td>
                </tr>
                <tr class="tdbg"> 
                  <td width="126" height="22"> <div align="right"><strong>大类名称：</strong></div></td>
                  <td width="218"> <input name="BigClassName" type="text" size="20" maxlength="30"> 
                  </td>
                </tr>
                <tr bgcolor="#C0C0C0" class="tdbg"> 
                  <td height="22" align="center">&nbsp; </td>
                  <td height="22" align="center"> <div align="left"> 
                      <input name="Action" type="hidden" id="Action" value="Add">
                      <input name="Add" type="submit" value=" 添 加 ">
                    </div></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table>
      <%
end if
%> </td>
  </tr>
</table>