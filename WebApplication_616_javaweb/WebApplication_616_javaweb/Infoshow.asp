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
.forecast ul li a {
 display: block; text-align:left; height:27px;
}
.forecast ul li a:link {
 text-decoration:none;
}
.forecast ul li a:visited {
text-decoration:none;
}
.forecast ul li a:hover {
text-decoration:none;font-weight:bold; height:27px;
}
</style>
</head>
<body>
<!--#include file="head.asp"-->
<%
dim flag
flag	=	trim(request.QueryString("flag"))
if not isnumeric(flag) then
	response.Write("传递参数出错")
	response.End()
end if

sql="select * from config where flag="&flag
set rs=conn.execute(sql)
if not rs.eof then
	flag	=	rs("flag")
	infoTitle		=	trim(rs("Title"))
	infocontent	=	trim(rs("Content"))
end if
rs.close
set rs=nothing
%>
<div class="path">
	<p>您的位置：<a href="default.asp"><%=SiteName%></a> > <%=infoTitle%></p>
</div>
<div class="basic btm10">
	<div class="left">
		<div class="zh_recommed" id="zh_recommed">
			<h2><span><%=infoTitle%></span></h2>
      <div class="info"><%=infocontent%></div>
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