<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim ArticleID,Action,sqlProperty,rsProperty,FoundErr,ErrMsg,PurviewChecked
dim page,bigid
ArticleID=trim(request("ArticleID"))
Action=Trim(Request("Action"))
page=trim(request("page"))
bigid=trim(request("bigid"))
if page<>"" then
	page=page
else
	page=1
end if
if bigid<>"" then
	bigid=bigid
else
	bigid=""
end if
FoundErr=False
if ArticleId="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>请先选定产品！</li>"
end if
if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	call SetProperty()
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect "Admin_ArticleManage.asp?page=" & page & "&bigid=" & bigid
else
	call CloseConn()
	call WriteErrMsg()
end if

sub SetProperty()
	if instr(ArticleID,",")>0 then
		ArticleID=replace(ArticleID," ","")
		sqlProperty="select OnTop,Elite,Passed,adddate from Article where ArticleID in (" & ArticleID & ")"
	else
		ArticleID=Clng(ArticleID)
		sqlProperty="select OnTop,Elite,Passed,adddate  from article where ArticleID=" & ArticleID
	end if
	Set rsProperty= Server.CreateObject("ADODB.Recordset")
	rsProperty.open sqlProperty,conn,1,3
	do while not rsProperty.eof
			select case Action
			case "SetOnTop"
				rsProperty("OnTop") = True
			case "CancelOnTop"
				rsProperty("OnTop") = False
			case "SetElite"
				rsProperty("Elite") = True
			case "CancelElite"
				rsProperty("Elite") = False
			case "SetPassed"
				rsProperty("Passed") =True
			case "CancelPassed"
				rsProPerty("Passed") =False
			case "aa"
				rsProPerty("adddate")=Now()
			end select
			rsProperty.update
		rsProperty.movenext
	loop
	rsProperty.close
	set rsProperty=nothing
end sub
%>