<%@language=vbscript codepage=936 %>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim sql,rsBigClass,rsSmallClass,ErrMsg
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * From QDClass",conn,1,1
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
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("������Ӵ������ƣ�");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("С�����Ʋ���Ϊ�գ�");
	document.form2.SmallClassName.focus();
	return false;
  }
}
function ConfirmDelBig()
{
   if(confirm("ȷ��Ҫɾ���˴�����ɾ���˴���ͬʱ��ɾ����������С��͸����µ��������ţ����Ҳ��ָܻ���"))
     return true;
   else
     return false;
	 
}

function ConfirmDelSmall()
{
   if(confirm("ȷ��Ҫɾ����С����ɾ����С��ͬʱ��ɾ�������µ��������ţ����Ҳ��ָܻ���"))
     return true;
   else
     return false;
	 
}
</script>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><table width="90" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="20" align="center" valign="top"><a href="Admin_NewsClassAddBig1.asp"><strong><font color="#FF0000"><u><br />
            ������ŷ���</u></font></strong></a><br />
            <br> 
            <table width="500" border="0" cellpadding="0" cellspacing="1" bgcolor="#666666">
              <tr class="title"> 
                <td width="50%" height="20" align="center" bgcolor="#A4B6D7"><strong>��Ŀ����</strong></td>
                <td align="center"><strong>����ѡ��</strong></td>
              </tr>
              <%
	do while not rsBigClass.eof
%>
              <tr bgcolor="#F2F8FF" class="tdbg"> 
                <td width="233" height="22" style="padding-left:10px;">
				<%=rsBigClass("ClassName")%>
				</td>
                <td align="right" style="padding-right:10">
				&nbsp;<a href="Admin_NewsClassModifyBig1.asp?BigClassID=<%=rsBigClass("ClassID")%>">�޸�</a> 
                  | <a href="Admin_NewsClassDelBig1.asp?BigClassName=<%=rsBigClass("ClassName")%>&BigClassID=<%=rsBigClass("ClassID")%>" onClick="return ConfirmDelBig();">ɾ��</a></td>
              </tr>
              <%
			  rsBigClass.movenext
	loop
%>
            </table>
          </td>
        </tr>
      </table>
      <%
rsBigClass.close
set rsBigClass=nothing
%> </td>
  </tr>
</table>