<%@language=vbscript codepage=936 %>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<%
BigClassID=trim(Request("BigClassID"))
Action=trim(Request("Action"))
NewBigClassName=trim(request("NewBigClassName"))

if BigClassID="" then
  response.Redirect("Admin_NewsClassManage1.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from QDClass where ClassID=" & CLng(BigClassID),conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�˲�Ʒ���಻���ڣ�</li>"
else
	if Action="Modify" then
		if NewBigClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��Ʒ����������Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("ClassName")=NewBigClassName
    	 	rs.update
			rs.Close
	     	set rs=Nothing	
			call CloseConn()
     		Response.Redirect "Admin_NewsClassManage1.asp" 
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
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.BigClassName.focus();
    return false;
  }
}
</script>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><table width="80%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center" valign="top"><br>
            <strong>�� �� �� �� �� ��</strong> <br> <br> <form name="form1" method="post" action="Admin_NewsClassModifyBig1.asp">
              <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
                <tr bgcolor="#A4B6D7" class="title"> 
                  <td height="20" colspan="2" align="center"><strong>�޸Ĵ�������</strong></td>
                </tr>
                <tr class="tdbg"> 
                  <td width="126"><strong>����ID��</strong></td>
                  <td width="204"><%=rs("ClassID")%> <input name="BigClassID" type="hidden" id="BigClassID" value="<%=rs("ClassID")%>"> 
                    </td>
                </tr>
                <tr class="tdbg"> 
                  <td width="126"><strong>�������ƣ�</strong></td>
                  <td> <input name="NewBigClassName" type="text" id="NewBigClassName" value="<%=rs("ClassName")%>" size="20" maxlength="30"></td>
                </tr>
                <tr class="tdbg"> 
                  <td align="center">&nbsp;</td>
                  <td align="center"> <div align="left"> 
                      <input name="Action" type="hidden" id="Action" value="Modify">
                      <input name="Save" type="submit" id="Save" value=" �� �� ">
                    </div></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table>
      <%
	end if
end if
rs.close
set rs=nothing
%> </td>
  </tr>
</table>