<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteTitle%></title>
<meta name="keywords" content="<%=Sitekeyword%>">
<meta name="description" content="<%=sitename%>">
<META NAME="robots" CONTENT="index, follow">
<META NAME="GOOGLEBOT" CONTENT="INDEX, FOLLOW">
<meta name="searchtitle" content="<%=SiteKeyword%>">
<meta NAME="author" CONTENT="<%=sitename%>">
<meta NAME="copyright" content="<%=sitename%>">
<link href="css/css.CSS" rel="stylesheet" type="text/css" />
</head>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<%
	NewsID=request.QueryString("ID")
	conn.execute("update news set hits=hits+1 where newsID="&NewsID)
	set rs=server.CreateObject("Adodb.recordset")
	sql="select * from news where newsID="&NewsID
	rs.open sql,conn,1,1
	if rs.eof then
		Call ShowErr("参数错误！","News.asp")
	else
		arttitle=rs("title")
		content=rs("content")
		classID=rs("ClassID")
		Source=rs("Source")
		Author=rs("author")
		hits=rs("hits")
		adddate=formatdatetime(rs("adddate"),2)
	end if
	rs.close

ClassID=request("ClassID")
if ClassID<>"" then
	ClassName=conn.execute("Select ClassName from NewsClass where ClassID="&ClassID)(0)
else
	ClassName=conn.execute("Select Top 1 ClassName from NewsClass order by ClassID")(0)
end if
%>
<body>
<script language=javascript id=clientEventHandlersJS>
<!--
var number=8;

function LMYC() {
var lbmc;
//var treePic;
    for (i=1;i<=number;i++) {
        lbmc = eval('LM' + i);
        //treePic = eval('treePic'+i);
        //treePic.src = 'pic/file.gif';
        lbmc.style.display = 'none';
    }
}
 
function ShowFLT(i) {
    lbmc = eval('LM' + i);
    //treePic = eval('treePic' + i)
    if (lbmc.style.display == 'none') {
        LMYC();
        //treePic.src = 'pic/nofile.gif';
        lbmc.style.display = '';
    }
    else {
        //treePic.src = 'pic/file.gif';
        lbmc.style.display = 'none';
    }
}
//-->
      </script>
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><!--#include file="top.asp" --></td>
  </tr>
  <tr>
    <td align="left" valign="top"><table width="1003" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="287" align="left" valign="top" bgcolor="#F6F6F6" class="index_td1"><!--#include file="left.asp"--></td>
        <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="37" background="pic/index_3.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="4%"><img src="pic/index_8.gif" width="24" height="37" /></td>
                <td width="22%" class="index_td9"><%=ClassName%></td>
                <td width="74%" align="right" style="padding-right:8px; padding-top:7px">您现在的位置：<a href="index.asp">首页</a>&gt;&gt;<%=ClassName%>&gt;&gt;正文</td>
              </tr>
            </table></td>
          </tr>
          
          <tr>
            <td align="left" valign="top"><div style="margin:5px 0; padding:5px; text-align:center">
                <h2 style="margin:0; font-size:18px"><%=arttitle%></h2>
            </div>
                <div style="padding:4px; width:97%; text-align:center; border-bottom:1px solid #ddd; border-right:1px solid #ddd; background:#f3f3f3; margin:0 auto;">来源：<%=source%> | 添加人：<%=author%> | 添加时间：<%=adddate%> | 点击次数：<%=hits%> | <a href="javascript:window.close()"><font color="#ff9900">关闭</font></a></div>
              <div id="newscontent" style="padding:10px; font-size:14px; line-height:180%; text-align:left; width:95%; margin:0 auto">
                  <%
			p=request.QueryString("p")
			str=content
			url="news_show.asp?id="&newsID&"&"
			set newpage=new aspxsky_page
			newpage.showpage str,p,url
			%>
              </div></td>
          </tr>
          
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include file="foot.asp" --></td>
  </tr>
</table>
</body>
</html>
