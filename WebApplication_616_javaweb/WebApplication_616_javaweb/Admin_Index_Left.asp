<!--#include file="inc/check.asp"--><html>
<head>
<title>管理导航</title>
<style type=text/css>
body  { background:#39867B; margin:0px; font:9pt 宋体; }
table  { border:0px; }
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; }
a:hover  { color:#cc0000;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#C6EBDE; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#39867B; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#cc0000; font-weight:bold; }
.STYLE1 {color: #C6EBDE}
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left>
    <tr><td valign=top>
<table width=158 border="0" align=center cellpadding=0 cellspacing=0>
  <tr>
    <td height=42 valign=bottom>
	  <img src="images/title.gif" width=158 height=38>
    </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/title_bg_quit.gif id=menuTitle0> 
          <span><a href="Admin_Index_Main.asp" target=main><b>管理首页</b></a> | 
			<a href=Admin_logout.asp target=_top><b>退出</b></a></span> 
        </td>
  </tr>
  <tr>
    <td style="display:" id='submenu0'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=150>
<tr><td height=20>用户名：<%=session("AdminName")%></td></tr>
<tr><td height=20>欢迎您：<%=session("PName")%></td></tr>
</table>
</div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
</div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(1)" style="cursor:hand;"> 
          <span>新闻管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu1'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20><a href=Admin_NewsAdd.asp target=main>添加新闻</a></td></tr>
<tr><td height=20><a href=Admin_NewsManage.asp target=main>新闻管理</a></td></tr>
<tr>
  <td height=20><a href="Admin_NewsClassManage.asp" target="main">新闻类别</a></td>
</tr>
<!--
<tr>
  <td height=20><a href="Admin_NewsClassManage.asp" target="main">类别管理</a></td>
</tr>
-->
</table>
</div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
</div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_9.gif" id=menuTitle2 onClick="showsubmenu(2)" style="cursor:hand;"> 
          <span>文章管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu2'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
  <td height=20><a href="Admin_ArticleAdd.asp" target="main">添加文章</a></td>
</tr>
<tr>
  <td height=20><a href="Admin_ArticleManage.asp" target="main">文章管理</a></td>
</tr>
<tr>
  <td height=20><a href="Admin_Class_Article.asp" target="main">文章分类</a></td>
</tr>
</table>
</div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
</div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_2.gif" id=menuTitle4 onClick="showsubmenu(4)" style="cursor:hand;"> 
          <span>常规设置</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu4'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td height=20><a href="Admin_SiteConfig.asp" target="main">配置信息</a> | <a href="Admin_ConfigManage.asp?flag=1" target="main">课程简介</a></td>
              </tr>
              <tr>
                <td height=20><a href="Admin_ConfigManage.asp?flag=2" target="main">申报书</a> | <a href="Admin_ConfigManage.asp?flag=3" target="main">政策支持</a></td>
				     </tr>
              <tr>
                <td height=20><a href="Admin_ConfigManage.asp?flag=4" target="main">课程特色</a> <!--| <a href="Admin_ConfigManage.asp?flag=5" target="main">营销网络</a> </td>
              </tr>
              <tr>
                <td height=20><a href="Admin_ConfigManage.asp?flag=6" target="main">售后服务</a> | <a href="Admin_ConfigManage.asp?flag=7" target="main">联系我们</a></td>
              </tr>-->
              <tr>
                <td height=20><a href="Admin_ConfigManage.asp?flag=8" target="main">底部信息</a></td>
              </tr>
              <tr>
                <td height=20><a href="/bbs" target="main">论坛管理</a></td>
              </tr>
			      <tr>
                <td height=20><a href="img.asp?action=showadd" target="main">图片添加</a> | <a href="img.asp" target="main">图片管理</a></td>
              </tr>
            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_4.gif" id=menuTitle7 onClick="showsubmenu(7)" style="cursor:hand;"> 
          <span>用户管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu7'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr>
                <td height=20><a href=Admin_Admin.asp?Action=Add target=main>管理员添加</a> 
                  | <a href=Admin_Admin.asp target=main>管理</a></td>
              </tr>
            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_9.gif" id=menuTitle8> 
          <span>技术支持</span> </td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=145>
<tr><td height=20>
    <a href="http://hi.baidu.com/bangbangnt" target="_blank">bangbangnt</a></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
</td></tr>
</table>
</body>
</html>
