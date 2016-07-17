<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="<%=Sitekeyword%>">
<meta name="description" content="<%=sitename%>">
<META NAME="robots" CONTENT="index, follow">
<META NAME="GOOGLEBOT" CONTENT="INDEX, FOLLOW">
<META NAME="baiduspider" CONTENT="INDEX, FOLLOW">
<META NAME="slurp" CONTENT="INDEX, FOLLOW">
<meta name="date" content="<%=now()%>" scheme="yyyy/mm/dd hh:mm:ss" />
<link href="images/style.css" rel="stylesheet" type="text/css" />
<title><%=SiteTitle%></title>
<style type="text/css">
.info ul { padding:2px 10px}
.info ul li { font-size:14px}
.info ul li span { float:right; text-align:left; color:#F60}
</style>
</head>
<body>
<!--#include file="head.asp"-->
<%
dim ClassID
ClassID		=	trim(request.QueryString("ClassID"))
if not isnumeric(ClassID) then
	response.Write("传递参数出错")
	response.End()
end if
bigname=ShowclassName(ClassID)
%>
<div class="path">
	<p>您的位置：<a href="default.asp"><%=SiteName%></a> > <%=bigname%></p>
</div>
<div class="basic btm10">
	<div class="left">
		<div class="living">
			<h2>最新公告</h2>
			<p id="srollingTips">
<%
	sql="select top 5 NewsID,Title,Adddatetime from news where ClassID=3 order by adddate desc"
	set rs=conn.execute(sql)
	do until rs.eof
	NewsID	=	rs("NewsID")
	Title		=	trim(rs("Title"))
	Adddatetime	=	formatdatetime(rs("Adddatetime"),2)
%>      
				<span><%=title%>&nbsp;&nbsp;发布时间：<em><%=Adddatetime%></em></span>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
			</p>
		</div>

		<div class="zh_recommed" id="zh_recommed">
			<h2><span><%=bigname%></span></h2>
      <div class="info">
<%
	sql="select newsid,title,Adddate from news where classid="&classid&" order by Adddate desc,newsid desc"
		set rs=server.CreateObject("adodb.recordset")
		'response.Write(sql)
		rs.open sql,conn,1,1
		if rs.eof then
			response.Write("暂时没有文章!")
		else
			response.Write("<ul>")
			CurrentPage = Request("page")
			if Not IsNumeric(CurrentPage) Then
				CurrentPage = "1"
			end if
			CurrentPage=Cint(CurrentPage)
			rs.PageSize = 20
			If CurrentPage < 1 Then CurrentPage = 1
			If CurrentPage > rs.PageCount Then 
				CurrentPage = rs.PageCount
			end if
			rs.AbsolutePage = CurrentPage
			count=rs.recordcount
			exh_i=0
		do while not rs.eof
		newsid	=	rs("newsid")		'新闻id
		title		=	gotTopic(rs("title"),60)			'标题
		Adddate	=	formatdatetime(rs("Adddate"),2)
%>
	<li><span><%=Adddate%></span>·<a href="news.asp?newsid=<%=newsid%>" target="_blank"><%=title%></a></li>
<%		
		exh_i=exh_i+1 
		if exh_i >= rs.PageSize Then exit do 
		rs.movenext
		loop
		response.Write("</ul>")
	end if
	call showpage("newlist.asp?classid="&classid,count,rs.pagesize,true,true,"篇文章")
	rs.close
	set rs=nothing
%>
      </div>
		</div>
    
	</div>
	<div class="right">
	<!--#include file="right.asp"-->
	</div>
	<div class="clear"></div>
</div>
<!--#include file="copyright.asp"-->
<script type="text/javascript" src="jquery.js"></script>
<script type="text/javascript" src="js.js"></script>
</body>
</html>