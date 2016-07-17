<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim ObjInstalled,Action,FoundErr,ErrMsg
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
Action=trim(request("Action"))
if Action="" then
	Action="ShowInfo"
end if
%>
<html>
<head>
<title>专题管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="topbg"> 
      
    <td height="22" colspan=2 align=center><b>网 站 配 置</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="70" height="30"><strong>管理导航：</strong></td>
      
    <td height="30"><a href="Admin_SiteConfig.asp">网站信息配置</a> </td>
    </tr>
  </table>
<%
if Action="SaveConfig" then
	call SaveConfig()
else
	call ShowConfig()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub ShowConfig()
%>
<form method="POST" action="Admin_SiteConfig.asp" id="form1" name="form1">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr> 
      <td height="22" colspan="2" class="topbg"> <a name="SiteInfo"></a><strong>网站信息配置</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>网站名称：</strong></td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteName" type="text" id="SiteName" value="<%=SiteName%>" size="40" maxlength="50">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>网站标题：</strong></td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteTitle" type="text" id="SiteTitle" value="<%=SiteTitle%>" size="40" maxlength="50">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>网站地址：</strong><br>
        请添写完整URL地址</td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteUrl" type="text" id="SiteUrl" value="<%=SiteUrl%>" size="40" maxlength="255">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>网站关键字</strong></td>
      <td height="25" class="tdbg"><input name="SiteKeyword" type="text" id="SiteKeyword" value="<%=SiteKeyword%>" size="40" maxlength="255"></td>
    </tr>
<!--    <tr>
      <td height="25" colspan="2" class="topbg"><strong>精彩视频开关</strong></td>
    </tr>
	<tr>
<td width="400" height="25" class="tdbg"><strong>开关：</strong></td>
      <td width="368" height="25" class="tdbg"><%
	  if vod="open" then
	  %><input type="radio" name="vod" value="open" checked="checked">
        开
          <input type="radio" name="vod" value="close">
      关
	  <%else%>
	  <input type="radio" name="vod" value="open">
        开
          <input type="radio" name="vod" value="close" checked="checked">
      关<%end if%></td>
	</tr>
    <tr>
      <td height="25" colspan="2" class="topbg"><strong>首页飘浮广告开关</strong></td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>开关：</strong></td>
      <td width="368" height="25" class="tdbg"><%
	  if ad_show="open" then
	  %><input type="radio" name="ad_show" value="open" checked="checked">
        开
          <input type="radio" name="ad_show" value="close">
      关
	  <%else%>
	  <input type="radio" name="ad_show" value="open">
        开
          <input type="radio" name="ad_show" value="close" checked="checked">
      关<%end if%></td>
    </tr>-->
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="UpFile" id="UpFile"></a><strong>上传文件选项</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>上传文件大小限制：</strong><br>
        建议不要超过1024K，以免影响服务器性能</td>
      <td height="25" class="tdbg">
        <input name="MaxFileSize" type="text" id="MaxFileSize" value="<%=MaxFileSize%>" size="6" maxlength="5">
        K</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>存放上传文件的目录：</strong><br>
        请输入相对于首页（Default.asp）的相对路径</td>
      <td height="25" class="tdbg">
        <input name="SaveUpFilesPath" type="text" id="SaveUpFilesPath" value="<%=SaveUpFilesPath%>" size="30" maxlength="100">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>允许的上传文件类型：</strong><br>
        只输入扩展名。每种文件类型用“|”号分开。</td>
      <td height="25" class="tdbg">
        <input name="UpFileType" type="text" id="UpFileType" value="<%=UpFileType%>" size="50" maxlength="255">      </td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"> 
        <input name="Action" type="hidden" id="Action" value="SaveConfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " <% If ObjInstalled=false Then response.write "disabled" %>>      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr class='tdbg'><td height='40' colspan='3'><b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能。<br>请直接修改“Inc/config.asp”文件中的内容。</font></b></td></tr>"
End If
%>
  </table>
<%
end sub
%>
</form>
</body>
</html>
<%
sub SaveConfig()
	If ObjInstalled=false Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>你的服务器不支持 FSO(Scripting.FileSystemObject)! </li>"
		exit sub
	end if
	dim fso,hf
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("/inc/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const SiteName=" & chr(34) & trim(request("SiteName")) & chr(34) & "        '网站名称" & vbcrlf
	hf.write "Const SiteTitle=" & chr(34) & trim(request("SiteTitle")) & chr(34) & "        '网站标题" & vbcrlf
	hf.write "Const SiteKeyword=" & chr(34) & trim(request("SiteKeyword")) & chr(34) & "        '网站关键字" & vbcrlf
	hf.write "Const SiteUrl=" & chr(34) & trim(request("SiteUrl")) & chr(34) & "        '网站地址" & vbcrlf
	hf.write "Const MaxFileSize=" & trim(request("MaxFileSize")) & "        '上传文件大小限制" & vbcrlf
	hf.write "Const vod=" & chr(34) & trim(request("vod")) & chr(34) & "        '视频开关" & vbcrlf
	hf.write "Const ad_show=" & chr(34) & trim(request("ad_show")) & chr(34) & "        '广告开关" & vbcrlf
	hf.write "Const SaveUpFilesPath=" & chr(34) & trim(request("SaveUpFilesPath")) & chr(34) & "        '存放上传文件的目录" & vbcrlf
	hf.write "Const UpFileType=" & chr(34) & trim(request("UpFileType")) & chr(34) & "        '允许的上传文件类型" & vbcrlf
	hf.write "%" & ">"
	hf.close
	set hf=nothing
	set fso=nothing
	call WriteSuccessMsg("网站配置保存成功！")
end sub
%>