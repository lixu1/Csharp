<%
'=================================================
'��������ShowGuestPage()
'��  �ã���ʾ���Եײ���ҳ
'��  ������
'=================================================
sub ShowGuestPage()
	dim PageFileName
	PageFileName=strFileName
	if keyword<>"" then
		PageFileName=PageFileName&"?keyword="&keyword
	end if
	if action<>"" then
		PageFileName=PageFileName&"?action="&action
	end if
	if TotalPut>0 then
		call showpage(PageFileName,totalPut,MaxPerPage,true,true,"������")
	end if
end sub

'=================================================
'��������SaveWriteGuest()
'��  �ã���������
'��  ������
'=================================================
sub SaveWriteGuest()
	dim SaveEditId,GuestName,GuestSex,GuestOicq,GuestEmail,GuestHomepage,GuestFace,GuestImages,GuestIcq,GuestMsn
	dim GuestTitle,GuestContent,GuestIsPrivate
	dim GuestPassword,GuestPasswordConfirm,GuestQuestion,GuestAnswer,GuestUserlevel
	dim sqlMaxId,rsMaxId,MaxId,Saveinfo,sqlReg,rsReg,CheckCode
	
	CheckCode=replace(trim(Request("CheckCode")),"'","")
	if CheckCode=""  then
		call Guest_info("<li>���Է���ʧ�ܣ�</li><li>��������֤�룡</li>")
		exit sub
	end if
	'if  CheckCode<>CStr(session("CheckCode")) then
	if  CheckCode<>CStr(session("getcode")) then
		call Guest_info("<li>���Է���ʧ�ܣ�</li><li>���������֤���ϵͳ�����Ĳ�һ�£�����������</li>")
		exit sub
	end if
	GuestContent=request("GuestContent")
	if GuestType=0 then
		GuestName=trim(request("GuestName"))
		GuestSex=trim(request("GuestSex"))
		GuestOicq=trim(request("GuestOicq"))
		GuestIcq=trim(request("GuestIcq"))
		GuestMsn=trim(request("GuestMsn"))
		GuestEmail=trim(request("GuestEmail"))
		GuestHomepage=trim(request("GuestHomepage"))
		if GuestHomepage="http://" or isnull(GuestHomepage) then GuestHomepage=""
	else
		GuestName=LoginName
	end if
	GuestImages=trim(request("GuestImages"))
	GuestFace=trim(request("GuestFace"))
	GuestTitle=trim(request("GuestTitle"))
	GuestIsPrivate=trim(request("GuestIsPrivate"))
	if GuestIsPrivate="yes" then
		GuestIsPrivate=True
	else
		GuestIsPrivate=False
	end if
	SaveEdit=request("saveedit")
		
	if GuestName="" or GuestTitle="" or GuestContent="" then
		call Guest_info("<li>���Է���ʧ�ܣ�</li><li>�뽫��Ҫ����Ϣ��д������</li>")
		exit sub
	end if

	if SaveEdit=1 then
		SaveEditId=request("saveeditid")
		if SaveEditId="" then
			call Guest_info("<li>��ָ��Ҫ�༭������ID��</li>")
			exit sub
		else
			sqlMaxId="select max(GuestMaxId) as MaxId from Guest"
			set rsMaxId=conn.execute(sqlMaxId)
			MaxId=rsMaxId("MaxId")
			if MaxId="" or isnull(MaxId) then MaxId=0
			set rsGuest=server.createobject("adodb.recordset")
			sql="select * from Guest where GuestID="&SaveEditId
			rsGuest.open sql,conn,1,3
			rsGuest("GuestName")=GuestName
			rsGuest("GuestSex")=GuestSex
			rsGuest("GuestOicq")=GuestOicq
			rsGuest("GuestIcq")=GuestIcq
			rsGuest("GuestMsn")=GuestMsn
			rsGuest("GuestEmail")=GuestEmail
			rsGuest("GuestHomepage")=GuestHomepage
			rsGuest("GuestIP")=Request.ServerVariables("REMOTE_ADDR")
			rsGuest("GuestTitle")=GuestTitle
			rsGuest("GuestFace")=GuestFace
			rsGuest("GuestContent")=GuestContent
			rsGuest("GuestDatetime")=now()
			rsGuest("GuestImages")=GuestImages
			rsGuest("GuestMaxId")=MaxId+1
			rsGuest("GuestIsPrivate")=GuestIsPrivate
			rsGuest.update
			call Guest_info("<li>���Ա༭�ɹ���</li><li>ֻ�й���Ա���ͨ�������ԲŻ���ʾ������</li>")
		end if
	else
		if GuestContent<>session("OldGuestContent") then 
			session("OldGuestContent")=GuestContent
			sqlMaxId="select max(GuestMaxId) as MaxId from Guest"
			set rsMaxId=conn.execute(sqlMaxId)
			MaxId=rsMaxId("MaxId")
			if MaxId="" or isnull(MaxId) then MaxId=0
			set rsGuest=server.createobject("adodb.recordset")
			sql="select * from Guest"
			rsGuest.open sql,conn,1,3
			rsGuest.addnew
			rsGuest("GuestType")=GuestType
			rsGuest("GuestName")=GuestName
			rsGuest("GuestSex")=GuestSex
			rsGuest("GuestOicq")=GuestOicq
			rsGuest("GuestIcq")=GuestIcq
			rsGuest("GuestMsn")=GuestMsn
			rsGuest("GuestEmail")=GuestEmail
			rsGuest("GuestHomepage")=GuestHomepage
			rsGuest("GuestIP")=Request.ServerVariables("REMOTE_ADDR")
			rsGuest("GuestTitle")=GuestTitle
			rsGuest("GuestFace")=GuestFace
			rsGuest("GuestContent")=GuestContent
			rsGuest("GuestDatetime")=now()
			rsGuest("GuestImages")=GuestImages
			rsGuest("GuestMaxId")=MaxId+1
			rsGuest("GuestIsPrivate")=GuestIsPrivate
			rsGuest.update
			Saveinfo="<li>���������Ѿ����ͳɹ���</li><li>ֻ�й���Ա���ͨ�������ԲŻ���ʾ������</li>"

			call Guest_info(Saveinfo)
		else
			call Guest_info("<li>�벻Ҫ��������������ͬ�����Ի�����ԣ�</li>")
		end if
	end if
end sub

'=================================================
'��������SendRegEmail()
'��  �ã�����ע��ȷ����
'��  ������
'=================================================
sub SendRegEmail()
	dim MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority
	MailtoAddress=Email
	MailtoName=RegUserName
	Subject="ע��ȷ����"
	MailBody="����һ��ע��ȷ���š������֤���ǣ�" & CheckNum & vbcrlf & "<br>���˽���ȷ�ϣ�<a href='" & CheckUrl & "'>" & CheckUrl & "</a>"
	FromName=SiteName
	MailFrom=WebmasterEmail
	Priority=3
	ErrMsg=SendMail(MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority)
	if ErrMsg<>"" then FoundErr=True
end sub


'=================================================
'��������Guest_info()
'��  �ã����Բ�����Ϣ
'��  ����info ��ʾ��Ϣ����
'=================================================
sub Guest_info(info)
	dim strInfo
	strInfo=strInfo & "<br><table cellpadding=2 cellspacing=1 border=0 width=460 class='border' align=center>" & vbcrlf
	strInfo=strInfo & "<tr align='center' class='title'><td><strong>���Բ���������Ϣ</strong></td></tr>" & vbcrlf
	strInfo=strInfo & "<tr class='tdbg'><td height='100' valign='top'><br>" & info &"</td></tr>" & vbcrlf
	strInfo=strInfo & "<tr align='center' class='tdbg'><td><a href='" & strFileName & "'>���鿴���ԡ�</a>" & vbcrlf
	strInfo=strInfo & "</td></tr></table><br>" & vbcrlf
	response.write strInfo
end sub


'=================================================
'��������sub ShowTip()
'��  �ã���ʾ��ʾ
'��  ������
'=================================================
sub ShowTip()
%>
<DIV id=toolTipLayer style="POSITION: absolute; VISIBILITY: hidden"></DIV>
<SCRIPT language=JavaScript>

var ns4 = document.layers;
var ns6 = document.getElementById && !document.all;
var ie4 = document.all;
offsetX = 0;
offsetY = 20;
var toolTipSTYLE="";
function initToolTips()
{
  if(ns4||ns6||ie4)
  {
    if(ns4) toolTipSTYLE = document.toolTipLayer;
    else if(ns6) toolTipSTYLE = document.getElementById("toolTipLayer").style;
    else if(ie4) toolTipSTYLE = document.all.toolTipLayer.style;
    if(ns4) document.captureEvents(Event.MOUSEMOVE);
    else
    {
      toolTipSTYLE.visibility = "visible";
      toolTipSTYLE.display = "none";
    }
    document.onmousemove = moveToMouseLoc;
  }
}
function toolTip(msg, fg, bg)
{
  if(toolTip.arguments.length < 1) // hide
  {
    if(ns4) toolTipSTYLE.visibility = "hidden";
    else toolTipSTYLE.display = "none";
  }
  else // show
  {
    if(!fg) fg = "#333333";
    if(!bg) bg = "#FFFFFF";
    var content =
    '<table border="0" cellspacing="0" cellpadding="1" bgcolor="' + fg + '"><td>' +
    '<table border="0" cellspacing="0" cellpadding="1" bgcolor="' + bg + 
    '"><td align="left" nowrap style="line-height: 120%"><font color="' + fg +
    '">' + msg +
    '&nbsp\;</font></td></table></td></table>';
    if(ns4)
    {
      toolTipSTYLE.document.write(content);
      toolTipSTYLE.document.close();
      toolTipSTYLE.visibility = "visible";
    }
    if(ns6)
    {
      document.getElementById("toolTipLayer").innerHTML = content;
      toolTipSTYLE.display='block'
    }
    if(ie4)
    {
      document.all("toolTipLayer").innerHTML=content;
      toolTipSTYLE.display='block'
    }
  }
}
function moveToMouseLoc(e)
{
  if(ns4||ns6)
  {
    x = e.pageX;
    y = e.pageY;
  }
  else
  {
    x = event.x + document.body.scrollLeft;
    y = event.y + document.body.scrollTop;
  }
  toolTipSTYLE.left = x + offsetX;
  toolTipSTYLE.top = y + offsetY;
  return true;
}
initToolTips();
</SCRIPT>
<%
end sub

'=================================================
'��������sub ShowUbb()
'��  �ã���ʾUbb��ǩ
'��  ����info ��ʾ��Ϣ����
'=================================================
sub ShowUbb()
%>
<script language=JavaScript>
helpstat = false;
stprompt = true;
basic = false;
function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 0) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 2) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}

function AddText(NewCode) {
document.formwrite.GuestContent.value+=NewCode;
}

function emails() {
	if (helpstat) {
		alert("Email ���\n���� Email ��������\n�÷�1: [email]nobody@domain.com[/email]\n�÷�2: [email=nobody@domain.com]����[/email]");
	} else if (basic) {
		AddTxt="[email][/email]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("������ʾ������.\n���Ϊ�գ���ô��ֻ��ʾ��� Email ��ַ",""); 
		if (txt2!=null) {
			txt=prompt("Email ��ַ.","name@domain.com");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[email]"+txt+"[/email]";
				} else {
					AddTxt="[email="+txt+"]"+txt2;
					AddText(AddTxt);
					AddTxt="[/email]";
				} 
				AddText(AddTxt);	        
			}
		}
	}
}

function flash() {
 	if (helpstat){
		alert("Flash ����\n���� Flash ����.\n�÷�: [flash]Flash �ļ��ĵ�ַ[/flash]");
	} else if (basic) {
		AddTxt="[flash][/flash]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("Flash �ļ��ĵ�ַ","http://");
		if (txt!=null) {             
			AddTxt="[flash]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
		}        
	}  
}

function Cdir() {
 	if (helpstat){
		alert("Shockwave ����\n���� Shockwave ����.\n�÷�: [dir=500,350]Shockwave �ļ��ĵ�ַ[/dir]");
	} else if (basic) {
		AddTxt="[dir][/dir]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("Shockwave �ļ��ĵ�ַ","");
		if (txt!=null) {             
			AddTxt="[dir=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/dir]";
			AddText(AddTxt);
		}        
	}  
}

function Crm() {
 	if (helpstat){
		alert("real player �ļ�\n���� real player �ļ�.\n�÷�: [rm=500,350]real player �ļ��ĵ�ַ[/rm]");
	} else if (basic) {
		AddTxt="[rm][/rm]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("real player �ļ��ĵ�ַ","");
		if (txt!=null) {             
			AddTxt="[rm=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/rm]";
			AddText(AddTxt);
		}        
	}  
}

function Cwmv() {
 	if (helpstat){
		alert("media player �ļ�\n���� wmv �ļ�.\n�÷�: [mp=500,350]media player �ļ��ĵ�ַ[/mp]");
	} else if (basic) {
		AddTxt="[mp][/mp]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("media player �ļ��ĵ�ַ","");
		if (txt!=null) {             
			AddTxt="[mp=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/mp]";
			AddText(AddTxt);
		}        
	}  
}

function Cmov() {
 	if (helpstat){
		alert("quick time �ļ�\n���� quick time �ļ�.\n�÷�: [qt=500,350]quick time �ļ��ĵ�ַ[/qt]");
	} else if (basic) {
		AddTxt="[qt][/qt]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("quick time �ļ��ĵ�ַ","");
		if (txt!=null) {             
			AddTxt="[qt=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/qt]";
			AddText(AddTxt);
		}        
	}  
}


function showsize(size) {
	if (helpstat) {
		alert("���ִ�С���\n�������ִ�С.\n�ɱ䷶Χ 1 - 6.\n 1 Ϊ��С 6 Ϊ���.\n�÷�: [size="+size+"]���� "+size+" ����[/size]");
	} else if (basic) {
		AddTxt="[size="+size+"][/size]";
		AddText(AddTxt);
	} else {                       
		txt=prompt("��С "+size,"����"); 
		if (txt!=null) {             
			AddTxt="[size="+size+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/size]";
			AddText(AddTxt);
		}        
	}
}

function bold() {
	if (helpstat) {
		alert("�Ӵֱ��\nʹ�ı��Ӵ�.\n�÷�: [b]���ǼӴֵ�����[/b]");
	} else if (basic) {
		AddTxt="[b][/b]";
		AddText(AddTxt);
	} else {  
		txt=prompt("���ֽ������.","����");     
		if (txt!=null) {           
			AddTxt="[b]"+txt;
			AddText(AddTxt);
			AddTxt="[/b]";
			AddText(AddTxt);
		}       
	}
}

function italicize() {
	if (helpstat) {
		alert("б����\nʹ�ı������Ϊб��.\n�÷�: [i]����б����[/i]");
	} else if (basic) {
		AddTxt="[i][/i]";
		AddText(AddTxt);
	} else {   
		txt=prompt("���ֽ���б��","����");     
		if (txt!=null) {           
			AddTxt="[i]"+txt;
			AddText(AddTxt);
			AddTxt="[/i]";
			AddText(AddTxt);
		}	        
	}
}

function quote() {
	if (helpstat){
		alert("���ñ��\n����һЩ����.\n�÷�: [quote]��������[/quote]");
	} else if (basic) {
		AddTxt="[quote][/quote]";
		AddText(AddTxt);
	} else {   
		txt=prompt("�����õ�����","����");     
		if(txt!=null) {          
			AddTxt="[quote]"+txt;
			AddText(AddTxt);
			AddTxt="[/quote]";
			AddText(AddTxt);
		}	        
	}
}

function showcolor(color) {
	if (helpstat) {
		alert("��ɫ���\n�����ı���ɫ.  �κ���ɫ�������Ա�ʹ��.\n�÷�: [color="+color+"]��ɫҪ�ı�Ϊ"+color+"������[/color]");
	} else if (basic) {
		AddTxt="[color="+color+"][/color]";
		AddText(AddTxt);
	} else {  
     	txt=prompt("ѡ�����ɫ��: "+color,"����");
		if(txt!=null) {
			AddTxt="[color="+color+"]"+txt;
			AddText(AddTxt);        
			AddTxt="[/color]";
			AddText(AddTxt);
		} 
	}
}

function center() {
 	if (helpstat) {
		alert("������\nʹ��������, ����ʹ�ı�����롢���С��Ҷ���.\n�÷�: [align=center|left|right]Ҫ������ı�[/align]");
	} else if (basic) {
		AddTxt="[align=center|left|right][/align]";
		AddText(AddTxt);
	} else {  
		txt2=prompt("������ʽ\n���� 'center' ��ʾ����, 'left' ��ʾ�����, 'right' ��ʾ�Ҷ���.","center");               
		while ((txt2!="") && (txt2!="center") && (txt2!="left") && (txt2!="right") && (txt2!=null)) {
			txt2=prompt("����!\n����ֻ������ 'center' �� 'left' ���� 'right'.","");               
		}
		txt=prompt("Ҫ������ı�","�ı�");     
		if (txt!=null) {          
			AddTxt="\r[align="+txt2+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/align]";
			AddText(AddTxt);
		}	       
	}
}

function hyperlink() {
	if (helpstat) {
		alert("�������ӱ��\n����һ���������ӱ��\nʹ�÷���: [url]http://www.asp163.net[/url]\nUSE: [url=http://www.asp163.net]��������[/url]");
	} else if (basic) {
		AddTxt="[url][/url]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("�����ı���ʾ.\n�������ʹ��, ����Ϊ��, ��ֻ��ʾ�������ӵ�ַ. ",""); 
		if (txt2!=null) {
			txt=prompt("��������.","http://");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url]"+txt;
					AddText(AddTxt);
					AddTxt="[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url="+txt+"]"+txt2;
					AddText(AddTxt);
					AddTxt="[/url]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function image() {
	if (helpstat){
		alert("ͼƬ���\n����ͼƬ\n�÷��� [img]http://www.asp163.net/logo.gif[/img]");
	} else if (basic) {
		AddTxt="[img][/img]";
		AddText(AddTxt);
	} else {  
		txt=prompt("ͼƬ�� URL","http://");    
		if(txt!=null) {            
			AddTxt="[img]"+txt;
			AddText(AddTxt);
			AddTxt="[/img]";
			AddText(AddTxt);
		}	
	}
}

function showcode() {
	if (helpstat) {
		alert("������\nʹ�ô����ǣ�����ʹ��ĳ����������� html �ȱ�־���ᱻ�ƻ�.\nʹ�÷���:\n [code]�����Ǵ�������[/code]");
	} else if (basic) {
		AddTxt="\r[code]\r[/code]";
		AddText(AddTxt);
	} else {   
		txt=prompt("�������","");     
		if (txt!=null) {          
			AddTxt="[code]"+txt;
			AddText(AddTxt);
			AddTxt="[/code]";
			AddText(AddTxt);
		}	       
	}
}

function list() {
	if (helpstat) {
		alert("�б���\n����һ�����ֻ��������б�.\n\nUSE: [list] [*]��Ŀһ[/*] [*]��Ŀ��[/*] [*]��Ŀ��[/*] [/list]");
	} else if (basic) {
		AddTxt=" [list][*]  [/*][*]  [/*][*]  [/*][/list]";
		AddText(AddTxt);
	} else {  
		txt=prompt("�б�����\n���� 'A' ��ʾ�����б�, '1' ��ʾ�����б�, ���ձ�ʾ�����б�.","");               
		while ((txt!="") && (txt!="A") && (txt!="a") && (txt!="1") && (txt!=null)) {
			txt=prompt("����!\n����ֻ������ 'A' �� '1' ��������.","");               
		}
		if (txt!=null) {
			if (txt=="") {
				AddTxt="[list]";
			} else {
				AddTxt="[list="+txt+"]";
			} 
			txt="1";
			while ((txt!="") && (txt!=null)) {
				txt=prompt("�б���\n�հױ�ʾ�����б�",""); 
				if (txt!="") {             
					AddTxt+="[*]"+txt+"[/*]"; 
				}                   
			} 
			AddTxt+="[/list] ";
			AddText(AddTxt); 
		}
	}
}

function showfont(font) {
 	if (helpstat){
		alert("������\n��������������.\n�÷�: [face="+font+"]�ı���������Ϊ"+font+"[/face]");
	} else if (basic) {
		AddTxt="[face="+font+"][/face]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("Ҫ�������������"+font,"����");
		if (txt!=null) {             
			AddTxt="[face="+font+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/face]";
			AddText(AddTxt);
		}        
	}  
}
function underline() {
  	if (helpstat) {
		alert("�»��߱��\n�����ּ��»���.\n�÷�: [u]Ҫ���»��ߵ�����[/u]");
	} else if (basic) {
		AddTxt="[u][/u]";
		AddText(AddTxt);
	} else {  
		txt=prompt("�»�������.","����");     
		if (txt!=null) {           
			AddTxt="[u]"+txt;
			AddText(AddTxt);
			AddTxt="[/u]";
			AddText(AddTxt);
		}	        
	}
}
function setfly() {
 	if (helpstat){
		alert("������\nʹ���ַ���.\n�÷�: [fly]����Ϊ��������[/fly]");
	} else if (basic) {
		AddTxt="[fly][/fly]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("��������","����");
		if (txt!=null) {             
			AddTxt="[fly]"+txt;
			AddText(AddTxt);
			AddTxt="[/fly]";
			AddText(AddTxt);
		}        
	}  
}

function move() {
	if (helpstat) {
		alert("�ƶ����\nʹ���ֲ����ƶ�Ч��.\n�÷�: [move]Ҫ�����ƶ�Ч��������[/move]");
	} else if (basic) {
		AddTxt="[move][/move]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Ҫ�����ƶ�Ч��������","����");     
		if (txt!=null) {           
			AddTxt="[move]"+txt;
			AddText(AddTxt);
			AddTxt="[/move]";
			AddText(AddTxt);
		}       
	}
}

function shadow() {
	if (helpstat) {
               alert("��Ӱ���\nʹ���ֲ�����ӰЧ��.\n�÷�: [SHADOW=���, ��ɫ, �߽�]Ҫ������ӰЧ��������[/SHADOW]");
	} else if (basic) {
		AddTxt="[SHADOW=255,blue,1][/SHADOW]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("���ֵĳ��ȡ���ɫ�ͱ߽��С","255,blue,1"); 
		if (txt2!=null) {
			txt=prompt("Ҫ������ӰЧ��������","����");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[SHADOW=255, blue, 1]"+txt;
					AddText(AddTxt);
					AddTxt="[/SHADOW]";
					AddText(AddTxt);
				} else {
					AddTxt="[SHADOW="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/SHADOW]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function glow() {
	if (helpstat) {
		alert("���α��\nʹ���ֲ�������Ч��.\n�÷�: [GLOW=���, ��ɫ, �߽�]Ҫ��������Ч��������[/GLOW]");
	} else if (basic) {
		AddTxt="[glow=255,red,2][/glow]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("���ֵĳ��ȡ���ɫ�ͱ߽��С","255,red,2"); 
		if (txt2!=null) {
			txt=prompt("Ҫ��������Ч��������.","����");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[glow=255,red,2]"+txt;
					AddText(AddTxt);
					AddTxt="[/glow]";
					AddText(AddTxt);
				} else {
					AddTxt="[glow="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/glow]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function openemot()
{
	var Win =window.open("guestselect.asp?action=emot","face","width=380,height=300,resizable=1,scrollbars=1");
}
function openhelp()
{
	var Win =window.open("editor_ubbhelp.asp","face","width=550,height=400,resizable=1,scrollbars=1");
}

</script>

<img onclick=bold() align=absmiddle src="<%=GuestPath%>Ubb_bold.gif" width="22" height="22" alt="����" border="0"
><img onclick=italicize() align=absmiddle src="<%=GuestPath%>Ubb_italicize.gif" width="23" height="22" alt="б��" border="0"
><img onclick=underline() align=absmiddle src="<%=GuestPath%>Ubb_underline.gif" width="23" height="22" alt="�»���" border="0"
><img onclick=center() align=absmiddle src="<%=GuestPath%>Ubb_center.gif" width="22" height="22" alt="����" border="0"
><img onclick=hyperlink() align=absmiddle src="<%=GuestPath%>Ubb_url.gif" width="22" height="22" alt="��������" border="0"
><img onclick=emails() align=absmiddle src="<%=GuestPath%>Ubb_email.gif" width="23" height="22" alt="Email����" border="0"
><img onclick=image() align=absmiddle src="<%=GuestPath%>Ubb_image.gif" width="23" height="22" alt="ͼƬ" border="0"
><img onclick=flash() align=absmiddle src="<%=GuestPath%>Ubb_swf.gif" width="23" height="22" alt="FlashͼƬ" border="0"
><img onclick=Cdir() align=absmiddle src="<%=GuestPath%>Ubb_Shockwave.gif" width="23" height="22" alt="Shockwave�ļ�" border="0"
><img onclick=Crm() align=absmiddle src="<%=GuestPath%>Ubb_rm.gif" width="23" height="22" alt="realplay��Ƶ�ļ�" border="0"
><img onclick=Cwmv() align=absmiddle src="<%=GuestPath%>Ubb_mp.gif" width="23" height="22" alt="Media Player��Ƶ�ļ�" border="0"
><img onclick=Cmov() align=absmiddle src="<%=GuestPath%>Ubb_qt.gif" width="23" height="22" alt="QuickTime��Ƶ�ļ�" border="0"
><img onclick=showcode() align=absmiddle src="<%=GuestPath%>Ubb_code.gif" width="22" height="22" alt="����" border="0"
><img onclick=quote() align=absmiddle src="<%=GuestPath%>Ubb_quote.gif" width="23" height="22" alt="����" border="0"
><IMG onclick=setfly() align=absmiddle height=22 alt=������ src="<%=GuestPath%>Ubb_fly.gif" width=23 border=0
><IMG onclick=move() align=absmiddle height=22 alt=�ƶ��� src="<%=GuestPath%>Ubb_move.gif" width=23 border=0
><IMG onclick=glow() align=absmiddle height=22 alt=������ src="<%=GuestPath%>Ubb_glow.gif" width=23 border=0
><IMG onclick=shadow() align=absmiddle height=22 alt=��Ӱ�� src="<%=GuestPath%>Ubb_shadow.gif" width=23 border=0
><img onclick=openhelp() align=absmiddle src="<%=GuestPath%>Ubb_Help.gif" width="23" height="22" alt="����" border="0"><br>
���壺
<SELECT onchange=showfont(this.options[this.selectedIndex].value) name=font>
	<option value="����" selected>����</option>
	<option value="����_GB2312">����</option>
	<option value="������">������</option>
	<option value="����">����</option>
	<option value="����">����</option>
	<OPTION value="Andale Mono">Andale Mono</OPTION> 
	<OPTION value=Arial>Arial</OPTION> 
	<OPTION value="Arial Black">Arial Black</OPTION> 
	<OPTION value="Book Antiqua">Book Antiqua</OPTION>
	<OPTION value="Century Gothic">Century Gothic</OPTION> 
	<OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
	<OPTION value="Courier New">Courier New</OPTION>
	<OPTION value=Georgia>Georgia</OPTION>
	<OPTION value=Impact>Impact</OPTION>
	<OPTION value=Tahoma>Tahoma</OPTION>
	<OPTION value="Times New Roman" >Times New Roman</OPTION>
	<OPTION value="Trebuchet MS">Trebuchet MS</OPTION>
	<OPTION value="Script MT Bold">Script MT Bold</OPTION>
	<OPTION value=Stencil>Stencil</OPTION>
	<OPTION value=Verdana>Verdana</OPTION>
	<OPTION value="Lucida Console">Lucida Console</OPTION>
</SELECT>
�����С
<select name="size" onChange="showsize(this.options[this.selectedIndex].value)">
	<option value="1">1</option>
	<option value="2">2</option>
	<option value="3" selected>3</option>
	<option value="4">4</option>
</select>
��ɫ��
<SELECT onchange=showcolor(this.options[this.selectedIndex].value) name=color> 
	<option style="background-color:#F0F8FF;color: #F0F8FF" value="#F0F8FF">#F0F8FF</option>
	<option style="background-color:#FAEBD7;color: #FAEBD7" value="#FAEBD7">#FAEBD7</option>
	<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">#00FFFF</option>
	<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4">#7FFFD4</option>
	<option style="background-color:#F0FFFF;color: #F0FFFF" value="#F0FFFF">#F0FFFF</option>
	<option style="background-color:#F5F5DC;color: #F5F5DC" value="#F5F5DC">#F5F5DC</option>
	<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4">#FFE4C4</option>
	<option style="background-color:#000000;color: #000000" value="#000000">#000000</option>
	<option style="background-color:#FFEBCD;color: #FFEBCD" value="#FFEBCD">#FFEBCD</option>
	<option style="background-color:#0000FF;color: #0000FF" value="#0000FF">#0000FF</option>
	<option style="background-color:#8A2BE2;color: #8A2BE2" value="#8A2BE2">#8A2BE2</option>
	<option style="background-color:#A52A2A;color: #A52A2A" value="#A52A2A">#A52A2A</option>
	<option style="background-color:#DEB887;color: #DEB887" value="#DEB887">#DEB887</option>
	<option style="background-color:#5F9EA0;color: #5F9EA0" value="#5F9EA0">#5F9EA0</option>
	<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">#7FFF00</option>
	<option style="background-color:#D2691E;color: #D2691E" value="#D2691E">#D2691E</option>
	<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">#FF7F50</option>
	<option style="background-color:#6495ED;color: #6495ED" value="#6495ED" selected>#6495ED</option>
	<option style="background-color:#FFF8DC;color: #FFF8DC" value="#FFF8DC">#FFF8DC</option>
	<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">#DC143C</option>
	<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">#00FFFF</option>
	<option style="background-color:#00008B;color: #00008B" value="#00008B">#00008B</option>
	<option style="background-color:#008B8B;color: #008B8B" value="#008B8B">#008B8B</option>
	<option style="background-color:#B8860B;color: #B8860B" value="#B8860B">#B8860B</option>
	<option style="background-color:#A9A9A9;color: #A9A9A9" value="#A9A9A9">#A9A9A9</option>
	<option style="background-color:#006400;color: #006400" value="#006400">#006400</option>
	<option style="background-color:#BDB76B;color: #BDB76B" value="#BDB76B">#BDB76B</option>
	<option style="background-color:#8B008B;color: #8B008B" value="#8B008B">#8B008B</option>
	<option style="background-color:#556B2F;color: #556B2F" value="#556B2F">#556B2F</option>
	<option style="background-color:#FF8C00;color: #FF8C00" value="#FF8C00">#FF8C00</option>
	<option style="background-color:#9932CC;color: #9932CC" value="#9932CC">#9932CC</option>
	<option style="background-color:#8B0000;color: #8B0000" value="#8B0000">#8B0000</option>
	<option style="background-color:#E9967A;color: #E9967A" value="#E9967A">#E9967A</option>
	<option style="background-color:#8FBC8F;color: #8FBC8F" value="#8FBC8F">#8FBC8F</option>
	<option style="background-color:#483D8B;color: #483D8B" value="#483D8B">#483D8B</option>
	<option style="background-color:#2F4F4F;color: #2F4F4F" value="#2F4F4F">#2F4F4F</option>
	<option style="background-color:#00CED1;color: #00CED1" value="#00CED1">#00CED1</option>
	<option style="background-color:#9400D3;color: #9400D3" value="#9400D3">#9400D3</option>
	<option style="background-color:#FF1493;color: #FF1493" value="#FF1493">#FF1493</option>
	<option style="background-color:#00BFFF;color: #00BFFF" value="#00BFFF">#00BFFF</option>
	<option style="background-color:#696969;color: #696969" value="#696969">#696969</option>
	<option style="background-color:#1E90FF;color: #1E90FF" value="#1E90FF">#1E90FF</option>
	<option style="background-color:#B22222;color: #B22222" value="#B22222">#B22222</option>
	<option style="background-color:#FFFAF0;color: #FFFAF0" value="#FFFAF0">#FFFAF0</option>
	<option style="background-color:#228B22;color: #228B22" value="#228B22">#228B22</option>
	<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">#FF00FF</option>
	<option style="background-color:#DCDCDC;color: #DCDCDC" value="#DCDCDC">#DCDCDC</option>
	<option style="background-color:#F8F8FF;color: #F8F8FF" value="#F8F8FF">#F8F8FF</option>
	<option style="background-color:#FFD700;color: #FFD700" value="#FFD700">#FFD700</option>
	<option style="background-color:#DAA520;color: #DAA520" value="#DAA520">#DAA520</option>
	<option style="background-color:#808080;color: #808080" value="#808080">#808080</option>
	<option style="background-color:#008000;color: #008000" value="#008000">#008000</option>
	<option style="background-color:#ADFF2F;color: #ADFF2F" value="#ADFF2F">#ADFF2F</option>
	<option style="background-color:#F0FFF0;color: #F0FFF0" value="#F0FFF0">#F0FFF0</option>
	<option style="background-color:#FF69B4;color: #FF69B4" value="#FF69B4">#FF69B4</option>
	<option style="background-color:#CD5C5C;color: #CD5C5C" value="#CD5C5C">#CD5C5C</option>
	<option style="background-color:#4B0082;color: #4B0082" value="#4B0082">#4B0082</option>
	<option style="background-color:#FFFFF0;color: #FFFFF0" value="#FFFFF0">#FFFFF0</option>
	<option style="background-color:#F0E68C;color: #F0E68C" value="#F0E68C">#F0E68C</option>
	<option style="background-color:#E6E6FA;color: #E6E6FA" value="#E6E6FA">#E6E6FA</option>
	<option style="background-color:#FFF0F5;color: #FFF0F5" value="#FFF0F5">#FFF0F5</option>
	<option style="background-color:#7CFC00;color: #7CFC00" value="#7CFC00">#7CFC00</option>
	<option style="background-color:#FFFACD;color: #FFFACD" value="#FFFACD">#FFFACD</option>
	<option style="background-color:#ADD8E6;color: #ADD8E6" value="#ADD8E6">#ADD8E6</option>
	<option style="background-color:#F08080;color: #F08080" value="#F08080">#F08080</option>
	<option style="background-color:#E0FFFF;color: #E0FFFF" value="#E0FFFF">#E0FFFF</option>
	<option style="background-color:#FAFAD2;color: #FAFAD2" value="#FAFAD2">#FAFAD2</option>
	<option style="background-color:#90EE90;color: #90EE90" value="#90EE90">#90EE90</option>
	<option style="background-color:#D3D3D3;color: #D3D3D3" value="#D3D3D3">#D3D3D3</option>
	<option style="background-color:#FFB6C1;color: #FFB6C1" value="#FFB6C1">#FFB6C1</option>
	<option style="background-color:#FFA07A;color: #FFA07A" value="#FFA07A">#FFA07A</option>
	<option style="background-color:#20B2AA;color: #20B2AA" value="#20B2AA">#20B2AA</option>
	<option style="background-color:#87CEFA;color: #87CEFA" value="#87CEFA">#87CEFA</option>
	<option style="background-color:#778899;color: #778899" value="#778899">#778899</option>
	<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE">#B0C4DE</option>
	<option style="background-color:#FFFFE0;color: #FFFFE0" value="#FFFFE0">#FFFFE0</option>
	<option style="background-color:#00FF00;color: #00FF00" value="#00FF00">#00FF00</option>
	<option style="background-color:#32CD32;color: #32CD32" value="#32CD32">#32CD32</option>
	<option style="background-color:#FAF0E6;color: #FAF0E6" value="#FAF0E6">#FAF0E6</option>
	<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">#FF00FF</option>
	<option style="background-color:#800000;color: #800000" value="#800000">#800000</option>
	<option style="background-color:#66CDAA;color: #66CDAA" value="#66CDAA">#66CDAA</option>
	<option style="background-color:#0000CD;color: #0000CD" value="#0000CD">#0000CD</option>
	<option style="background-color:#BA55D3;color: #BA55D3" value="#BA55D3">#BA55D3</option>
	<option style="background-color:#9370DB;color: #9370DB" value="#9370DB">#9370DB</option>
	<option style="background-color:#3CB371;color: #3CB371" value="#3CB371">#3CB371</option>
	<option style="background-color:#7B68EE;color: #7B68EE" value="#7B68EE">#7B68EE</option>
	<option style="background-color:#00FA9A;color: #00FA9A" value="#00FA9A">#00FA9A</option>
	<option style="background-color:#48D1CC;color: #48D1CC" value="#48D1CC">#48D1CC</option>
	<option style="background-color:#C71585;color: #C71585" value="#C71585">#C71585</option>
	<option style="background-color:#191970;color: #191970" value="#191970">#191970</option>
	<option style="background-color:#F5FFFA;color: #F5FFFA" value="#F5FFFA">#F5FFFA</option>
	<option style="background-color:#FFE4E1;color: #FFE4E1" value="#FFE4E1">#FFE4E1</option>
	<option style="background-color:#FFE4B5;color: #FFE4B5" value="#FFE4B5">#FFE4B5</option>
	<option style="background-color:#FFDEAD;color: #FFDEAD" value="#FFDEAD">#FFDEAD</option>
	<option style="background-color:#000080;color: #000080" value="#000080">#000080</option>
	<option style="background-color:#FDF5E6;color: #FDF5E6" value="#FDF5E6">#FDF5E6</option>
	<option style="background-color:#808000;color: #808000" value="#808000">#808000</option>
	<option style="background-color:#6B8E23;color: #6B8E23" value="#6B8E23">#6B8E23</option>
	<option style="background-color:#FFA500;color: #FFA500" value="#FFA500">#FFA500</option>
	<option style="background-color:#FF4500;color: #FF4500" value="#FF4500">#FF4500</option>
	<option style="background-color:#DA70D6;color: #DA70D6" value="#DA70D6">#DA70D6</option>
	<option style="background-color:#EEE8AA;color: #EEE8AA" value="#EEE8AA">#EEE8AA</option>
	<option style="background-color:#98FB98;color: #98FB98" value="#98FB98">#98FB98</option>
	<option style="background-color:#AFEEEE;color: #AFEEEE" value="#AFEEEE">#AFEEEE</option>
	<option style="background-color:#DB7093;color: #DB7093" value="#DB7093">#DB7093</option>
	<option style="background-color:#FFEFD5;color: #FFEFD5" value="#FFEFD5">#FFEFD5</option>
	<option style="background-color:#FFDAB9;color: #FFDAB9" value="#FFDAB9">#FFDAB9</option>
	<option style="background-color:#CD853F;color: #CD853F" value="#CD853F">#CD853F</option>
	<option style="background-color:#FFC0CB;color: #FFC0CB" value="#FFC0CB">#FFC0CB</option>
	<option style="background-color:#DDA0DD;color: #DDA0DD" value="#DDA0DD">#DDA0DD</option>
	<option style="background-color:#B0E0E6;color: #B0E0E6" value="#B0E0E6">#B0E0E6</option>
	<option style="background-color:#800080;color: #800080" value="#800080">#800080</option>
	<option style="background-color:#FF0000;color: #FF0000" value="#FF0000">#FF0000</option>
	<option style="background-color:#BC8F8F;color: #BC8F8F" value="#BC8F8F">#BC8F8F</option>
	<option style="background-color:#4169E1;color: #4169E1" value="#4169E1">#4169E1</option>
	<option style="background-color:#8B4513;color: #8B4513" value="#8B4513">#8B4513</option>
	<option style="background-color:#FA8072;color: #FA8072" value="#FA8072">#FA8072</option>
	<option style="background-color:#F4A460;color: #F4A460" value="#F4A460">#F4A460</option>
	<option style="background-color:#2E8B57;color: #2E8B57" value="#2E8B57">#2E8B57</option>
	<option style="background-color:#FFF5EE;color: #FFF5EE" value="#FFF5EE">#FFF5EE</option>
	<option style="background-color:#A0522D;color: #A0522D" value="#A0522D">#A0522D</option>
	<option style="background-color:#C0C0C0;color: #C0C0C0" value="#C0C0C0">#C0C0C0</option>
	<option style="background-color:#87CEEB;color: #87CEEB" value="#87CEEB">#87CEEB</option>
	<option style="background-color:#6A5ACD;color: #6A5ACD" value="#6A5ACD">#6A5ACD</option>
	<option style="background-color:#708090;color: #708090" value="#708090">#708090</option>
	<option style="background-color:#FFFAFA;color: #FFFAFA" value="#FFFAFA">#FFFAFA</option>
	<option style="background-color:#00FF7F;color: #00FF7F" value="#00FF7F">#00FF7F</option>
	<option style="background-color:#4682B4;color: #4682B4" value="#4682B4">#4682B4</option>
	<option style="background-color:#D2B48C;color: #D2B48C" value="#D2B48C">#D2B48C</option>
	<option style="background-color:#008080;color: #008080" value="#008080">#008080</option>
	<option style="background-color:#D8BFD8;color: #D8BFD8" value="#D8BFD8">#D8BFD8</option>
	<option style="background-color:#FF6347;color: #FF6347" value="#FF6347">#FF6347</option>
	<option style="background-color:#40E0D0;color: #40E0D0" value="#40E0D0">#40E0D0</option>
	<option style="background-color:#EE82EE;color: #EE82EE" value="#EE82EE">#EE82EE</option>
	<option style="background-color:#F5DEB3;color: #F5DEB3" value="#F5DEB3">#F5DEB3</option>
	<option style="background-color:#FFFFFF;color: #FFFFFF" value="#FFFFFF">#FFFFFF</option>
	<option style="background-color:#F5F5F5;color: #F5F5F5" value="#F5F5F5">#F5F5F5</option>
	<option style="background-color:#FFFF00;color: #FFFF00" value="#FFFF00">#FFFF00</option>
	<option style="background-color:#9ACD32;color: #9ACD32" value="#9ACD32">#9ACD32</option>
</SELECT>
&nbsp;

<%end sub%>