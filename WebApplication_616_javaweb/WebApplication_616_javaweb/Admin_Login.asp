<%@language=vbscript codepage=936 %>
<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'主要是使随机出现的图片数字随机
%>
<html>
<head>
<title>管理员登录</title>
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
		alert("请输入用户名！");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("请输入密码！");
		document.Login.Password.focus();
		return false;
	}
}
</script>
</head>
<body class="bgcolor" onLoad="SetFocus();">
<p>　</p>
<form name="Login" action="Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
    
  <table width="585" border="0" align="center" cellpadding="0" cellspacing="0" >
    <tr> 
      <td width="280" rowspan="2"><img src="Images/entry1.gif" width="280" height="246"> 
      </td>
      <td width="344" background="Images/entry2.gif"> <table width="100%" border="0" cellspacing="8" cellpadding="0" align="center">
          <tr align="center"> 
            <td height="38" colspan="2"><font color="#000000" size="3"><strong>管理员登录</strong></font> 
            </td>
          </tr>
          <tr> 
            <td align="right"><font color="#000000">用户名称：</font></td>
            <td><input name="UserName"  type="text"  maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onmouseover="this.style.background='#E1F4EE';" onmouseout="this.style.background='#E1F4EE'" onFocus="this.select(); "></td>
          </tr>
          <tr> 
            <td align="right"><font color="#000000">用户密码：</font></td>
            <td><input name="Password"  type="password" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onmouseover="this.style.background='#E1F4EE';" onmouseout="this.style.background='#E1F4EE'" onFocus="this.select(); "></td>
          </tr>
		  
          <tr> 
            <td colspan="2"> <div align="center"> 
                <input   type="submit" name="Submit" value=" 确认 " style="font-size: 9pt; height: 19; width: 60; ">
                &nbsp; 
                <input name="reset" type="reset"  id="reset" value=" 清除 " style="font-size: 9pt; height: 19; width: 60; ">
                <br>
              </div></td>
          </tr>
        </table></td>
    </tr>
    <tr><td height="3"></td></tr>
  </table>
  <p align="center">后台管理页面需要屏幕分辨率为 <font color="#FF0000"><strong>1024*768</strong></font> 
    或以上才能达到最佳浏览效果！<br>
    需要浏览器为<strong><font color="#FF0000"> </font></strong><font color="#FF0000"><strong>IE5.5</strong></font> 
    或以上版本才能正常运行！！！</p>
</form>
</body>
</html>
