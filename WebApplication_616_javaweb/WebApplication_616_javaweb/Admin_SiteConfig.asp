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
<title>ר�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="topbg"> 
      
    <td height="22" colspan=2 align=center><b>�� վ �� ��</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="70" height="30"><strong>��������</strong></td>
      
    <td height="30"><a href="Admin_SiteConfig.asp">��վ��Ϣ����</a> </td>
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
      <td height="22" colspan="2" class="topbg"> <a name="SiteInfo"></a><strong>��վ��Ϣ����</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>��վ���ƣ�</strong></td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteName" type="text" id="SiteName" value="<%=SiteName%>" size="40" maxlength="50">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>��վ���⣺</strong></td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteTitle" type="text" id="SiteTitle" value="<%=SiteTitle%>" size="40" maxlength="50">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>��վ��ַ��</strong><br>
        ����д����URL��ַ</td>
      <td width="368" height="25" class="tdbg"> 
        <input name="SiteUrl" type="text" id="SiteUrl" value="<%=SiteUrl%>" size="40" maxlength="255">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>��վ�ؼ���</strong></td>
      <td height="25" class="tdbg"><input name="SiteKeyword" type="text" id="SiteKeyword" value="<%=SiteKeyword%>" size="40" maxlength="255"></td>
    </tr>
<!--    <tr>
      <td height="25" colspan="2" class="topbg"><strong>������Ƶ����</strong></td>
    </tr>
	<tr>
<td width="400" height="25" class="tdbg"><strong>���أ�</strong></td>
      <td width="368" height="25" class="tdbg"><%
	  if vod="open" then
	  %><input type="radio" name="vod" value="open" checked="checked">
        ��
          <input type="radio" name="vod" value="close">
      ��
	  <%else%>
	  <input type="radio" name="vod" value="open">
        ��
          <input type="radio" name="vod" value="close" checked="checked">
      ��<%end if%></td>
	</tr>
    <tr>
      <td height="25" colspan="2" class="topbg"><strong>��ҳƮ����濪��</strong></td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>���أ�</strong></td>
      <td width="368" height="25" class="tdbg"><%
	  if ad_show="open" then
	  %><input type="radio" name="ad_show" value="open" checked="checked">
        ��
          <input type="radio" name="ad_show" value="close">
      ��
	  <%else%>
	  <input type="radio" name="ad_show" value="open">
        ��
          <input type="radio" name="ad_show" value="close" checked="checked">
      ��<%end if%></td>
    </tr>-->
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="UpFile" id="UpFile"></a><strong>�ϴ��ļ�ѡ��</strong></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>�ϴ��ļ���С���ƣ�</strong><br>
        ���鲻Ҫ����1024K������Ӱ�����������</td>
      <td height="25" class="tdbg">
        <input name="MaxFileSize" type="text" id="MaxFileSize" value="<%=MaxFileSize%>" size="6" maxlength="5">
        K</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>����ϴ��ļ���Ŀ¼��</strong><br>
        �������������ҳ��Default.asp�������·��</td>
      <td height="25" class="tdbg">
        <input name="SaveUpFilesPath" type="text" id="SaveUpFilesPath" value="<%=SaveUpFilesPath%>" size="30" maxlength="100">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>������ϴ��ļ����ͣ�</strong><br>
        ֻ������չ����ÿ���ļ������á�|���ŷֿ���</td>
      <td height="25" class="tdbg">
        <input name="UpFileType" type="text" id="UpFileType" value="<%=UpFileType%>" size="50" maxlength="255">      </td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"> 
        <input name="Action" type="hidden" id="Action" value="SaveConfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " <% If ObjInstalled=false Then response.write "disabled" %>>      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr class='tdbg'><td height='40' colspan='3'><b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ����ܡ�<br>��ֱ���޸ġ�Inc/config.asp���ļ��е����ݡ�</font></b></td></tr>"
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
		ErrMsg=ErrMsg & "<br><li>��ķ�������֧�� FSO(Scripting.FileSystemObject)! </li>"
		exit sub
	end if
	dim fso,hf
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("/inc/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const SiteName=" & chr(34) & trim(request("SiteName")) & chr(34) & "        '��վ����" & vbcrlf
	hf.write "Const SiteTitle=" & chr(34) & trim(request("SiteTitle")) & chr(34) & "        '��վ����" & vbcrlf
	hf.write "Const SiteKeyword=" & chr(34) & trim(request("SiteKeyword")) & chr(34) & "        '��վ�ؼ���" & vbcrlf
	hf.write "Const SiteUrl=" & chr(34) & trim(request("SiteUrl")) & chr(34) & "        '��վ��ַ" & vbcrlf
	hf.write "Const MaxFileSize=" & trim(request("MaxFileSize")) & "        '�ϴ��ļ���С����" & vbcrlf
	hf.write "Const vod=" & chr(34) & trim(request("vod")) & chr(34) & "        '��Ƶ����" & vbcrlf
	hf.write "Const ad_show=" & chr(34) & trim(request("ad_show")) & chr(34) & "        '��濪��" & vbcrlf
	hf.write "Const SaveUpFilesPath=" & chr(34) & trim(request("SaveUpFilesPath")) & chr(34) & "        '����ϴ��ļ���Ŀ¼" & vbcrlf
	hf.write "Const UpFileType=" & chr(34) & trim(request("UpFileType")) & chr(34) & "        '������ϴ��ļ�����" & vbcrlf
	hf.write "%" & ">"
	hf.close
	set hf=nothing
	set fso=nothing
	call WriteSuccessMsg("��վ���ñ���ɹ���")
end sub
%>