<!--#include file="inc/check.asp"--><html>
<head>
<title>������</title>
<style type=text/css>
body  { background:#39867B; margin:0px; font:9pt ����; }
table  { border:0px; }
td  { font:normal 12px ����; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ����; color:#000000; text-decoration:none; }
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
          <span><a href="Admin_Index_Main.asp" target=main><b>������ҳ</b></a> | 
			<a href=Admin_logout.asp target=_top><b>�˳�</b></a></span> 
        </td>
  </tr>
  <tr>
    <td style="display:" id='submenu0'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=150>
<tr><td height=20>�û�����<%=session("AdminName")%></td></tr>
<tr><td height=20>��ӭ����<%=session("PName")%></td></tr>
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
          <span>���Ź���</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu1'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20><a href=Admin_NewsAdd.asp target=main>�������</a></td></tr>
<tr><td height=20><a href=Admin_NewsManage.asp target=main>���Ź���</a></td></tr>
<tr>
  <td height=20><a href="Admin_NewsClassManage.asp" target="main">�������</a></td>
</tr>
<!--
<tr>
  <td height=20><a href="Admin_NewsClassManage.asp" target="main">������</a></td>
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
          <span>���¹���</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu2'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
  <td height=20><a href="Admin_ArticleAdd.asp" target="main">�������</a></td>
</tr>
<tr>
  <td height=20><a href="Admin_ArticleManage.asp" target="main">���¹���</a></td>
</tr>
<tr>
  <td height=20><a href="Admin_Class_Article.asp" target="main">���·���</a></td>
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
          <span>��������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu4'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td height=20><a href="Admin_SiteConfig.asp" target="main">������Ϣ</a> | <a href="Admin_ConfigManage.asp?flag=1" target="main">�γ̼��</a></td>
              </tr>
              <tr>
                <td height=20><a href="Admin_ConfigManage.asp?flag=2" target="main">�걨��</a> | <a href="Admin_ConfigManage.asp?flag=3" target="main">����֧��</a></td>
				     </tr>
              <tr>
                <td height=20><a href="Admin_ConfigManage.asp?flag=4" target="main">�γ���ɫ</a> <!--| <a href="Admin_ConfigManage.asp?flag=5" target="main">Ӫ������</a> </td>
              </tr>
              <tr>
                <td height=20><a href="Admin_ConfigManage.asp?flag=6" target="main">�ۺ����</a> | <a href="Admin_ConfigManage.asp?flag=7" target="main">��ϵ����</a></td>
              </tr>-->
              <tr>
                <td height=20><a href="Admin_ConfigManage.asp?flag=8" target="main">�ײ���Ϣ</a></td>
              </tr>
              <tr>
                <td height=20><a href="/bbs" target="main">��̳����</a></td>
              </tr>
			      <tr>
                <td height=20><a href="img.asp?action=showadd" target="main">ͼƬ���</a> | <a href="img.asp" target="main">ͼƬ����</a></td>
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
          <span>�û�����</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu7'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr>
                <td height=20><a href=Admin_Admin.asp?Action=Add target=main>����Ա���</a> 
                  | <a href=Admin_Admin.asp target=main>����</a></td>
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
          <span>����֧��</span> </td>
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
