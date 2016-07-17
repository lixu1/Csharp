<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim Action,rs,sql,ErrMsg,newsid
dim ClassID,Title,Content,source,author

Action=trim(request("Action"))
ClassID=trim(request.form("ClassID"))
Title=trim(request.form("Title"))
Content=trim(request.form("Content"))
source=trim(request.form("source"))
author=trim(request.form("author"))
newsid=trim(request.form("newsid"))
hits=trim(request.Form("hits"))
if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足</li>"
else
	call SaveArticle()
end if
if founderr=true then
	call WriteErrMsg()
end if

call CloseConn()


sub SaveArticle()
	if ClassID="" then
		founderr=true
		errmsg=errmsg & "<br><li>未指定文章所属栏目。</li>"
	else
		ClassID=CLng(ClassID)
	end if
	if Title="" then
		founderr=true
		errmsg=ErrMsg & "<br><li>文章标题不能为空</li>"
	end if
	if Content="" then
		founderr=true
		errmsg=errmsg & "<br><li>文章内容不能为空</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	
	set rs=server.createobject("adodb.recordset")
	if Action="Add" then
		sql="select top 1 * from QDNews" 
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()
		rs.update
		rs.close
		call CloseConn()
		response.redirect "Admin_NewsManage1.asp"
	end if
	if Action="Modify" then
  		if newsID="" then
			founderr=true
			errmsg=errmsg & "<br><li>不能确定新闻ID的值</li>"
		else
			newsID=Clng(newsID)
			sql="select * from QDnews where newsid=" & newsID
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
		call CloseConn()
		response.redirect "Admin_NewsManage1.asp"
	end if
end sub

sub SaveData()
	rs("ClassID")=ClassID
	rs("Title")=Title
	rs("Content")=Content
	Rs("source")=source
	Rs("Author")=author
	Rs("Hits")=hits
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

%>