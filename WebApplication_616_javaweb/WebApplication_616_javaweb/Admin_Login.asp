<%@language=vbscript codepage=936 %>
<%
option explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'��Ҫ��ʹ������ֵ�ͼƬ�������
%>
<html>
<head>
<title>����Ա��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
<script language=javascript>
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("�������û�����");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("���������룡");
		document.Login.Password.focus();
		return false;
	}
}
</script>
</head>
<body class="bgcolor" onLoad="SetFocus();">
<p>��</p>
<form name="Login" action="Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
    
  <table width="585" border="0" align="center" cellpadding="0" cellspacing="0" >
    <tr> 
      <td width="280" rowspan="2"><img src="Images/entry1.gif" width="280" height="246"> 
      </td>
      <td width="344" background="Images/entry2.gif"> <table width="100%" border="0" cellspacing="8" cellpadding="0" align="center">
          <tr align="center"> 
            <td height="38" colspan="2"><font color="#000000" size="3"><strong>����Ա��¼</strong></font> 
            </td>
          </tr>
          <tr> 
            <td align="right"><font color="#000000">�û����ƣ�</font></td>
            <td><input name="UserName"  type="text"  maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onmouseover="this.style.background='#E1F4EE';" onmouseout="this.style.background='#E1F4EE'" onFocus="this.select(); "></td>
          </tr>
          <tr> 
            <td align="right"><font color="#000000">�û����룺</font></td>
            <td><input name="Password"  type="password" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onmouseover="this.style.background='#E1F4EE';" onmouseout="this.style.background='#E1F4EE'" onFocus="this.select(); "></td>
          </tr>
		  
          <tr> 
            <td colspan="2"> <div align="center"> 
                <input   type="submit" name="Submit" value=" ȷ�� " style="font-size: 9pt; height: 19; width: 60; ">
                &nbsp; 
                <input name="reset" type="reset"  id="reset" value=" ��� " style="font-size: 9pt; height: 19; width: 60; ">
                <br>
              </div></td>
          </tr>
        </table></td>
    </tr>
    <tr><td height="3"></td></tr>
  </table>
  <p align="center">��̨����ҳ����Ҫ��Ļ�ֱ���Ϊ <font color="#FF0000"><strong>1024*768</strong></font> 
    �����ϲ��ܴﵽ������Ч����<br>
    ��Ҫ�����Ϊ<strong><font color="#FF0000"> </font></strong><font color="#FF0000"><strong>IE5.5</strong></font> 
    �����ϰ汾�����������У�����</p>
</form>
</body>
</html>