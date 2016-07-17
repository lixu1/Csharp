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
<style type="text/css">
<!--
#img {　　max-width: 570px;
　　width:expression(this.width > 570? "570px" : auto);
　　}
-->
</style>
</head>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<body>
<%
	ArticleID=request.QueryString("ID")
	set rs=server.CreateObject("Adodb.recordset")
	sql="select * from article where ArticleID="&ArticleID
	rs.open sql,conn,1,1
	if rs.eof then
		Call ShowErr("参数错误！","product.asp")
	else
		arttitle=rs("title")
		content=rs("content")
		classID=rs("BigID")
		SmallID=rs("SmallID")
		picurl=rs("defaultpic")
		adddate=formatdatetime(rs("adddate"),2)
							set rs1=server.CreateObject("adodb.recordset")
							if classID>0 then
							sql1="select * from BigClass where BigID="&classID
							rs1.open sql1,conn,1,1
							BigName=rs1("BigName")
							rs1.close
							end if
							if SmallID>0 then
							sql1="select * from smallclass where SmallID="&SmallID
							rs1.open sql1,conn,1,1
							SmallName=rs1("smallname")
							rs1.close
							end if
							set rs1=nothing
	end if
	rs.close
	className=conn.execute("select bigName from BigClass where BigID="&ClassID)(0)
%>
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
                <td width="22%" class="index_td9">产品信息</td>
                <td width="74%" align="right" style="padding-right:8px; padding-top:7px">您现在的位置：<a href="index.asp">首页</a>&gt;&gt;<a href="ProductInfo.asp">产品信息</a>
				<%
					if len(bigname)>0 then
						if len(smallname)>0 then
								%>
                        <a href=productinfo.asp?bigID=<%=classID%>><font color="#333">>><%=bigname%></font></a><a href=productinfo.asp?smallid=<%=SmallID%>><font color="#333">>><%=smallname%></font></a>
                        <%else%>
                        <a href=productinfo.asp?bigID=<%=classID%>><font color="#333">>><%=bigname%></font></a>
                      <%
						end if
					end if
							%></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td class="index_5"><table width="100%" border="0" cellspacing="8" cellpadding="8">
              <tr>
                <td align="center" valign="top"><div style="margin:5px 0; padding:5px; text-align:center; overflow:hidden">
                    <h2 style="margin:0; font-size:18px"><%=arttitle%></h2>
                </div>
                    <div style="padding:4px; width:97%; text-align:right; border-bottom:1px solid #ddd; border-right:1px solid #ddd; background:#f3f3f3;"> 
                      <table width="200" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="75"><a href="javascript:this.history.go(-1)"><font color="#FF6633">返回</font></a></td>
                          <td width="125"><a href="javascript:this.history.go(-1)"> </a></td>
                        </tr>
                      </table>
                      <a href="javascript:this.history.go(-1)"></a></div>
                  <div style="padding:10px; font-size:14px; line-height:180%; text-align:left; width:90%;">
                    <div style="padding:10px 0; text-align:center"><img src="<%=picurl%>" id="img" style="padding:1px; border:1px solid #aaa;" /></div>
                    <%
			p=request.QueryString("p")
			str=content
			url="product_show.asp?id="&ArticleID&"&"
			set newpage=new aspxsky_page
			newpage.showpage str,p,url
			%>
                  </div><a href="javascript:this.history.go(-1)"><font color="#FF6633">返回</font></a></td>
              </tr>
            </table></td>
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
