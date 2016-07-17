<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim Action,rs,sql,ErrMsg,FoundErr,ObjInstalled,sqlStr
dim ArticleID,BigID,SmallID,Title,Content,Hits
dim IncludePic,DefaultPicUrl,Passed,OnTop,Elite
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=false
Action=trim(request("Action"))
ArticleID=Trim(Request.Form("ArticleID"))
BigID=trim(request.form("BigClass"))
SmallID=trim(request.form("SmallClass"))
Title=trim(request.form("Title"))
Content=trim(request.form("Content"))
IncludePic=trim(request.form("IncludePic"))
DefaultPicUrl=trim(request.form("Defaultpic"))
OnTop=trim(request.form("OnTop"))
Elite=trim(request.form("Elite"))
Hits=trim(request.form("Hits"))
if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足</li>"
else
	call SaveArticle()
end if
if founderr=true then
	call WriteErrMsg()
else
	call SaveSuccess()
end if
call CloseConn()


sub SaveArticle()
	dim PurviewChecked
	if BigID="" then
		founderr=true
		errmsg=errmsg & "<br><li>未指定文章所属栏目</li>"
	else
		BigID=CLng(BigID)
	end if
	if Title="" then
		founderr=true
		errmsg=ErrMsg & "<br><li>文章名称不能为空</li>"
	end if
	if Content="" then
		founderr=true
		errmsg=errmsg & "<br><li>文章内容不能为空</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	dim strSiteUrl
	strSiteUrl=request.ServerVariables("HTTP_REFERER")
	strSiteUrl=lcase(left(strSiteUrl,instrrev(strSiteUrl,"/")))
	if Hits<>"" then
		Hits=CLng(Hits)
	else
		Hits=0
	end if
	
	set rs=server.createobject("adodb.recordset")
	if Action="Add" then
		sql="select top 1 * from Article" 
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()
		rs.update
		ArticleID=rs("ArticleID")
		rs.close
	elseif Action="Modify" then
  		if ArticleID="" then
			founderr=true
			errmsg=errmsg & "<br><li>不能确定ArticleID的值</li>"
		else
			ArticleID=Clng(ArticleID)
			sql="select * from article where articleid=" & ArticleID
			rs.open sql,conn,1,3
			if rs.bof and rs.eof then
				founderr=true
				errmsg=errmsg & "<br><li>找不到此文章，可能已经被其他人删除。</li>"
 			else
				call SaveData()
				rs.update
				rs.close
			end if
		end if
	else
		FoundErr=True
		ErrMsg="<br><li>参数错误!</li>"
	end if
	set rs=nothing
end sub

sub SaveData()
	rs("BigID")=BigID
	if SmallID<>"" and SmallID<>"无小类别" then rs("SmallID")=SmallID
	rs("Title")=Title
	rs("Content")=Content
	rs("Hits")=Hits
	if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
	if OnTop="yes" then
		rs("OnTop")=True
	else
		rs("OnTop")=False
	end if
	if Elite="yes" then
		rs("Elite")=True
	else
		rs("Elite")=False
	end if
	rs("DefaultPic")=DefaultPicUrl
end sub
	
sub SaveSuccess()
%>
<html>
<head>
<title></title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<br><br>
<table class="border" align=center width="400" border="0" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr align=center> 
    <td  height="22" align="center" class="title"><b> 
      <%if  Action="Add" then%>
      添加 
      <%else%>
      修改 
      <%end if%>
      信息成功</b></td>
  </tr>
  <tr> 
    <td><table width="100%" border="0" cellpadding="2" cellspacing="1">
        
        <tr class="tdbg"> 
          <td width="100" align="right"><strong>名称：</strong></td>
          <td><%= Title %></td>
        </tr>
      </table></td>
  </tr>
  <tr class="tdbg"> 
    <td height="30" align="center">【<a href="Admin_ArticleModify.asp?ArticleID=<%=ArticleID%>">修改此信息</a>】&nbsp;【<a href="Admin_ArticleAdd.asp">继续添加信息</a>】&nbsp;【<a href="Admin_ArticleManage.asp">信息管理</a>】</td>
  </tr>
</table>
</body>
</html>
<%
end sub


'==================================================
'过程名：ReplaceRemoteUrl
'作  用：替换字符串中的远程文件为本地文件并保存远程文件
'参  数：strContent ------ 要替换的字符串
'==================================================
function ReplaceRemoteUrl(strContent)
	dim re,RemoteFile,RemoteFileurl,SaveFilePath,SaveFileName,SaveFileType,arrSaveFileName,ranNum
	SaveFilePath = "UploadFiles"			'文件保存的本地路径
	if right(SaveFilePath,1)<>"/" then SaveFilePath=SaveFilePath&"/"
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp)))"
	Set RemoteFile = re.Execute(strContent)
	For Each RemoteFileurl in RemoteFile
		arrSaveFileName = split(RemoteFileurl,".")
		SaveFileType=arrSaveFileName(ubound(arrSaveFileName))
		ranNum=int(900*rnd)+100
		SaveFileName = SaveFilePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&SaveFileType	
		call SaveRemoteFile(SaveFileName,RemoteFileurl)
		strContent=Replace(strContent,RemoteFileurl,SaveFileName)
		if UploadFiles="" then
			UploadFiles=SaveFileName
		else
			UploadFiles=UploadFiles & "|" & SaveFileName
		end if
	Next
	ReplaceRemoteUrl=strContent
end function



'==================================================
'过程名：SaveRemoteFile
'作  用：保存远程的文件到本地
'参  数：LocalFileName ------ 本地文件名
'		 RemoteFileUrl ------ 远程文件URL
'==================================================
sub SaveRemoteFile(LocalFileName,RemoteFileUrl)
	dim Ads,Retrieval,GetRemoteData
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", RemoteFileUrl, False, "", ""
		.Send
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	Set Ads = Server.CreateObject("Adodb.Stream")
	With Ads
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile server.MapPath(LocalFileName),2
		.Cancel()
		.Close()
	End With
	Set Ads=nothing
end sub

'==================================================
'过程名：Admin_ShowPath
'作  用：显示类别
'参  数：Big   ------ 大类ID
'		 Small ------ 小类ID
'==================================================
sub Admin_ShowPath(Big)
	
	set Rs=server.CreateObject("adodb.recordset")
	sqlStr="select BigName from BigClass Where BigID=" & clng(Big)
	Rs.open sqlStr,conn,1,1
	if not Rs.eof then
		response.Write Rs("BigName")
	end if
	Rs.close
	
end sub
%>