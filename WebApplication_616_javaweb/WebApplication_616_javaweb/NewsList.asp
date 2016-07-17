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
ClassID=request("ClassID")
if ClassID<>"" then
	ClassName=conn.execute("Select ClassName from NewsClass where ClassID="&ClassID)(0)
else
	set ClassInfo=conn.execute("Select Top 1 ClassName,ClassID from NewsClass order by ClassID")
	ClassName=ClassInfo(0)
	ClassID=ClassInfo(1)
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
                <td width="74%" align="right" style="padding-right:8px; padding-top:7px">您现在的位置：<a href="index.asp">首页</a>&gt;&gt;<%=ClassName%></td>
              </tr>
            </table></td>
          </tr>
          
          <tr>
            <td class="index_5"><%
		set Rs=server.CreateObject("adodb.recordset")
		sqlStr="select hotpic,newsid,title,adddatetime from news where ClassID="&ClassID&" order by adddate desc, newsid desc"
		Rs.open sqlStr,conn,1,1
		Rs.pagesize=15
		page =clng(Request.QueryString("page"))
		if page<1 then page=1
		if page>Rs.pagecount then page=Rs.pagecount
		if not Rs.eof then
		rs.AbsolutePage = page
%>
	<ul class="art_list">
<%
		for i=1 to Rs.pagesize
		%>
                  <li><span><font color="#ED4308">[<%=formatdatetime(Rs("adddatetime"),2)%>]</font></a></span>· <a href="newsView.asp?id=<%=Rs("newsid")%>" target="_blank"><%=Rs("title")%></a></li>
                  <%
			rs.movenext
			if rs.eof then exit for
		next
		%>
		</ul>
		<%
		else
			response.Write("没有新闻")
		end if
	%><div style="padding:10px; text-align:center">
                  <%
			if ClassID<>"" then
				filename="&ClassID="&ClassID
			else
				filename=""
			end if
			if Rs.RecordCount<>0 then
				if rs.pagecount=1 then
					response.Write "上一页&nbsp;&nbsp;"
					response.Write "下一页&nbsp;&nbsp;"
				else
					if page=1 then
						response.Write "上一页&nbsp;&nbsp;"
						response.Write "<a href=?page="& (page+1) &filename&">下一页</a>&nbsp;&nbsp;"
					elseif page=Rs.pagecount then
						response.Write "<a href=?page=" & (page-1) &filename&">上一页</a>&nbsp;&nbsp;"
						response.Write "下一页&nbsp;&nbsp;"
					else
						response.Write "<a href=?page=" & (page-1) &filename&">上一页</a>&nbsp;&nbsp;"
						response.Write "<a href=?page="& (page+1) &filename&">下一页</a>&nbsp;&nbsp;"
					end if
				end if
			end if
			%>
&nbsp;&nbsp;当前页<%=page%>/<%=rs.pagecount%>，&nbsp;共<%=rs.RecordCount%>条信息 &nbsp;<%=rs.pagesize%>条每页
<%Rs.close%>
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
