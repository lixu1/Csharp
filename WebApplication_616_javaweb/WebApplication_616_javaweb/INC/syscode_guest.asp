<%
'option explicit
response.buffer=true	
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="conn.asp"-->
<!--#include file="ubbcode.asp"-->
<!--#include file="function.asp"-->
<!--#include file="admin_code_guest.asp"-->

<%
dim strChannel,sqlChannel,rsChannel,ChannelUrl,ChannelName
dim strFileName,MaxPerPage,totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime,founderr, errmsg,i
dim rs,sql,rsGuest,sqlGuest
dim PageTitle,strPath,strPageTitle
dim SkinID,ClassID,AnnounceCount
dim UserGuestName,UserType,UserSex,UserEmail,UserHomepage,UserOicq,UserIcq,UserMsn
dim WriteName,WriteType,WriteSex,WriteEmail,WriteOicq,WriteIcq,WriteMsn,WriteHomepage
dim WriteFace,WriteImages,WriteTitle,WriteContent,SaveEdit,SaveEditId
dim GuestType,LoginName,AdminReplyContent
dim Action,SubmitType,GuestPath,TitleName,keyword
Set rsGuest= Server.CreateObject("ADODB.Recordset")


strFileName="Message.asp"
GuestPath="images/guestbook/"
MaxPerPage=2
SaveEdit=0							
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
TitleName=ChannelName
Action=request("Action")
select case Action
	case "write"
		PageTitle="ǩд����"
	case "savewrite"
		PageTitle="��������"
	case "reply"
		PageTitle="�ظ�����"
	case "edit"
		PageTitle="�༭����"
	case "del"
		PageTitle="ɾ������"
	case else
		PageTitle="��վ����"
end select

GuestType=0

keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=ReplaceBadChar(keyword)
	keyword=Replace(keyword,"[","")
	keyword=Replace(keyword,"]","")
end if
if keyword<>"" then TitleName="�������� <font color=red>"&keyword&"</font> ������"


'=================================================
'��������GuestBook_Left()
'��  �ã���ʾ������Թ���
'��  ������
'=================================================
sub GuestBook_Left()
	set rsGuest=conn.execute("select count(*) from Guest where GuestIsPassed=False")
	if GuestType=1 then
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����ģʽ��<font color=green>�û�ģʽ</font>" & vbcrlf
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������ԣ�<font color=red>"&rsGuest(0)&"</font> ��" & vbcrlf
		response.write "<div align='center'><a href='" & strFileName & "?action=user' >���ҵ����ԡ�</a></div>"
	else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����ģʽ��<font color=green>�ο�ģʽ</font>" & vbcrlf
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������ԣ�<font color=red>"&rsGuest(0)&"</font> ��" & vbcrlf
	end if
	response.write "<div align='center'><a href='" & strFileName & "' >���鿴���ԡ�</a></div>" & vbcrlf
	response.write "<div align='center'><a href='" & strFileName & "?action=write' >��ǩд���ԡ�</a></div><br>" & vbcrlf
end sub

'=================================================
'��������GuestBook_Search()
'��  �ã���ʾ��������
'��  ������
'=================================================
sub GuestBook_Search()
	response.write "<table border='0' cellpadding='0' cellspacing='0'>"
	response.write "<form method=""post"" name='SearchForm' action='" & strFileName & "'>"
	response.write "<tr><td height='40' >"
	response.write "&nbsp;&nbsp;<input type='text' name='keyword'  size='15' value='�ؼ���' maxlength='50' onFocus='this.select();'>&nbsp;"
	response.write "<input type='submit' name='Submit'  value='����'>"
	response.write "</td></tr></form></table>"
end sub
%>

<%

'=================================================
'��������GuestBook()
'��  �ã����Ա����ܵ���
'��  ������
'=================================================
sub GuestBook()
	select case Action
		case "write"
			call WriteGuest()
		case "savewrite"
			call SaveWriteGuest()
		case "reply"
			call ReplyGuest()
		case "edit"
			call EditGuest()
		case "del"
			call DelGuest()
		case "user"
			call ShowAllGuest(3)
		case else
			call GuestMain()
	end select
end sub

'=================================================
'��������GuestMain()
'��  �ã�����������
'��  ������
'=================================================
sub GuestMain()
	if GuestType=1 then
		ShowAllGuest(1)
	else
		ShowAllGuest(2)
	end if
end sub

'=================================================
'��������ShowAllGuest()
'��  �ã���ҳ��ʾ��������
'��  ����ShowType-----  0Ϊ��ʾ����
'						1Ϊ��ʾ��ͨ����˼��û��Լ����������
'						2Ϊ��ʾ��ͨ����˵����ԣ������ο���ʾ��
'						3Ϊ��ʾ�û��Լ����������
'=================================================
sub ShowAllGuest(ShowType)
	if ShowType=1 then
		sqlGuest="select * from Guest where (GuestIsPassed=True or GuestName='"&LoginName&"')"
	elseif ShowType=2 then
		sqlGuest="select * from Guest where GuestIsPassed=True"
	elseif ShowType=3 then
		sqlGuest="select * from Guest where GuestName='"&LoginName&"'"
	elseif ShowType=4 then
		sqlGuest="select * from Guest where GuestIsPassed=False"
	else
		if keyword<>"" then
			sqlGuest="select * from Guest where 1"
		else
			sqlGuest="select * from Guest"
		end if
	end if
	if keyword<>"" then
		sqlGuest=sqlGuest & " and (GuestTitle like '%" & keyword & "%' or GuestContent like '%" & keyword & "%' or GuestName like '%" & keyword & "%' or GuestReply like '%" & keyword & "%') "
	end if

	sqlGuest=sqlGuest&" order by GuestMaxId desc"
	set rsGuest=server.createobject("adodb.recordset")
	rsGuest.open sqlGuest,conn,1,1
	if rsGuest.bof and rsGuest.eof then
		totalput=0
		response.write "<br><li>û���κ�����</li>"
	else
		totalput=rsGuest.recordcount
		if currentPage=1 then
			call ShowGuestList()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsGuest.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsGuest.bookmark
            	call ShowGuestList()
        	else
	        	currentPage=1
           		call ShowGuestList()
	    	end if
		end if
	end if
	rsGuest.close
	set rsGuest=nothing
end sub

'=================================================
'��������ShowGuestList()
'��  �ã���ʾ����
'��  ������
'=================================================
sub ShowGuestList()
	dim i,GuestTip,TipName,TipSex,TipEmail,TipOicq,TipHomepage,isdelUser
	i=0
	do while not rsGuest.eof
		isdelUser=0
		if rsGuest("GuestType")=1 then
			sql="select * from " & db_User_Table & " where " & db_User_Name & "='" & rsGuest("GuestName")&"'"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,Conn_User,1,1
			if not rs.bof and not rs.eof then
				UserGuestName=rs(db_User_Name)
				UserSex=rs(db_User_Sex)
				UserEmail=rs(db_User_Email)
				UserOicq=rs(db_User_QQ)
				UserIcq=rs(db_User_Icq)
				UserMsn=rs(db_User_Msn)
				UserHomepage=rs(db_User_Homepage)
			else
				isdelUser=1
			end if
		end if
		if rsGuest("GuestType")<>1 or isdelUser=1 then
			UserGuestName=rsGuest("GuestName")
			UserSex=rsGuest("GuestSex")
			UserEmail=rsGuest("GuestEmail")
			UserOicq=rsGuest("GuestOicq")
			UserIcq=rsGuest("GuestIcq")
			UserMsn=rsGuest("GuestMsn")
			UserHomepage=rsGuest("GuestHomepage")
		end if
		TipName=UserGuestName
		if isdelUser=1 then TipName=TipName&"����ɾ����"
		if TipEmail="" or isnull(TipEmail) then TipEmail="δ��"
		if TipOicq="" or isnull(TipOicq) then TipOicq="δ��"
		if TipHomepage="" or isnull(TipHomepage) then TipHomepage="δ��"
		if UserIcq="" or isnull(UserIcq) then UserIcq="δ��"
		if UserMsn="" or isnull(UserMsn) then UserMsn="δ��"
		if UserSex=1 then
			TipSex="����磩"
		elseif UserSex=0 then
			TipSex="(����)"
		else
			TipSex=""
		end if
		GuestTip="&nbsp;������"&TipName&" "&TipSex&"<br>&nbsp;�ʼ���"&TipEmail&"<br>&nbsp;QQ/msn��"&TipOicq&"<br>&nbsp;��ҳ��"&TipHomepage&"<br>&nbsp;ʱ�䣺"&rsGuest("GuestDatetime")
		%>
		<SCRIPT language=javascript>
		function CheckAll(form)
		{
		  for (var i=0;i<form.elements.length;i++)
			{
			var e = form.elements[i];
			if (e.Name != "chkAll"&&e.disabled!=true)
			   e.checked = form.chkAll.checked;
			}
		}
		</script>
		<table width="99%" border="0" cellspacing="0" cellpadding="0" align="center">
		  <tr>
			<td> 
			  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="border" style="border: 1px dotted #808080">
				<tr> 
				<tr > 
				  <td align="center" valign="top"> 
					<table width="100%" border="0" cellspacing="2" cellpadding="1"  style="border-bottom-style: solid; border-bottom-width: 1px;border-bottom-color:#B4C9E7">
					  <tr> 
						<td ><font color=green>&nbsp;&nbsp;����</font>:&nbsp;<%=(rsGuest("GuestTitle"))%></td>
						<td width="228"> <img src="<%=GuestPath%>posttime.gif" width="11" height="11" align="absmiddle"> 
						  <font color="#006633">�� 
						  <% =rsGuest("GuestDatetime")%>
						  </font> </td>
					  </tr>
					</table>
				  </td>
				</tr>
				<tr class="tdbg_leftall"> 
				  <td align="center" height="153" valign="top"> 
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
					  <%if rsGuest("GuestIsPassed")=True then%>
						<tr class="tdbg_leftall"> 
					  <%else%>
						<tr bgcolor="#f0f0f0"> 
					  <%end if%>
						<td width="100" align="center" height="130" valign="top"> 
						  <table width="100%" border="0" cellspacing="0" cellpadding="3" align="center" >
							<tr> 
							  <td valign="middle" align="center" width="100%" style="padding-top:10px;">
							    <%
							    response.write "<b>�û���Ϣ</b>"
							    %>
								<img src="<%=GuestPath%><%=rsGuest("GuestImages")%>.gif" width="80" height="90" onMouseOut=toolTip() onMouseOver="toolTip('<%=GuestTip%>')"><br><br>
								<%
									response.write  (UserGuestName)
								%>
							  </td>
							</tr>
						  </table>
						</td>
						<td align="center" height="153" width="1" bgcolor="#B4C9E7"> 
						</td>
						<td> 
						  <table width="100%" border="0" cellpadding="6" cellspacing="0" class="saytext" height="125" style="TABLE-LAYOUT: fixed">
							<tr> 
							  <td align="left" valign="top"><img src="<%=GuestPath%>face<%=rsGuest("GuestFace")%>.gif" width="19" height="19"> 
								<%
								  if rsGuest("GuestIsPrivate")=true then
									response.write "<font color=green>[����]</font>&nbsp;"
								  end if
								  response.write (ubbcode(dvHTMLEncode(rsGuest("GuestContent"))))
								  %>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" valign="bottom"> 
                        <%call ShowGuestreply()%>
                      </td>
                    </tr>
                  </table>
						  <table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#B4C9E7">
							<tr> 
							  <td></td>
							</tr>
						  </table>
						  <%call ShowGuestButton()%>
						</td>
					  </tr>
					</table>
				  </td>
				</tr>
			  </table>
			  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr> 
				  <td  height="15" align="center" valign="top"> 
					<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
					  <tr> 
						<td height="13" class="tdbg_left2"></td>
					  </tr>
					</table>
				  </td>
				</tr>
			  </table>
			  </td>
		  </tr>
		</table>
		<%
		rsGuest.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
end sub

'=================================================
'��������ShowGuestreply()
'��  �ã���ʾ�ظ�����
'��  ������
'=================================================
sub ShowGuestreply()
	if len(rsGuest("GuestReply")) >0 then
	%>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
  <tr> 
		<td> 
		  
      <table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr> 
          <td height="1" bgcolor="#B4C9E7"></td>
        </tr>
        <tr>
          <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="TABLE-LAYOUT: fixed">
              <tr> 
                <td><font color="#006633"> ����Ա<font color="#FF0000">[<%=rsGuest("GuestReplyAdmin")%>]</font>�ظ�: 
                  <% =rsGuest("GuestReplyDatetime") %>
                  </font> </td>
              </tr>
              <tr> 
                <td valign="bottom"><font color="#006633"><% =KeywordReplace(ubbcode(dvHTMLEncode(rsGuest("GuestReply")))) %></font></td>
              </tr>
            </table></td>
        </tr>
      </table>
		</td>
	  </tr>
	</table>
	<%
	end if
end sub

'**************************************************
'��������KeywordReplace
'��  �ã���ʾ�����ؼ���
'��  ����strChar-----Ҫת�����ַ�
'����ֵ��ת������ַ�
'**************************************************
function KeywordReplace(strChar)
	if strChar="" then
		KeywordReplace=""
	else
		KeywordReplace=	replace(strChar,""&keyword&"","<font color=red>"&keyword&"</font>")
	end if
end function

'=================================================
'��������WriteGuest()
'��  �ã�ǩд����
'��  ������
'=================================================
sub WriteGuest()
if SaveEdit<>1 then
	WriteType=GuestType
	WriteName=LoginName
	WriteSex="1"
	WriteFace="1"
	WriteImages="01"
	WriteHomepage="http://"
end if
%>
<script language=JavaScript>
function changeimage()
{ 
	document.formwrite.GuestImages.value=document.formwrite.Image.value;
	document.formwrite.showimages.src="<%=GuestPath%>"+document.formwrite.Image.value+".gif";
}
function check(thisform)
{
   if(thisform.GuestName.value==""){
		alert("��������Ϊ�գ�")
		 thisform.GuestName.focus()
		  return(false) 
	  }

   if(thisform.GuestTitle.value==""){
		alert("�������ⲻ��Ϊ�գ�")
		thisform.GuestTitle.focus()
		return(false)
	  }
   if(thisform.GuestIcq.value==""){
		alert("��ϵ�绰����Ϊ�գ�")
		thisform.GuestIcq.focus()
		return(false)
	  }
   if(thisform.GuestMsn.value==""){
		alert("��ϵ��ַ����Ϊ�գ�")
		thisform.GuestMsn.focus()
		return(false)
	  }
   if(thisform.GuestContent.value==""){
		alert("�������ݲ���Ϊ�գ�")
		thisform.GuestContent.focus()
		  return(false)
	  }
   if(thisform.CheckCode.value==""){
		alert("��֤�벻��Ϊ�գ�")
		thisform.CheckCode.focus()
		  return(false)
	  }

   if(thisform.GuestContent.value.length>500){
		alert("�������ݲ��ܳ���500�ַ���")
		thisform.GuestContent.focus()
		  return(false)
	  }

}
function showreginfo(){
if (document.formwrite.reg.checked == true) {
	reginfo.style.display = "";
//	reginfoshowtext.innerText="��ʱ����ע���Ϊ��վ��Ա"
}else{
	reginfo.style.display = "none";
//	reginfoshowtext.innerText="����ͬʱע���Ϊ��վ��Ա"
}
}
function gbcount(message,total,used,remain)
{
	var max;
	max = total.value;
	if (message.value.length > max) {
	message.value = message.value.substring(0,max);
	used.value = max;
	remain.value = 0;
	alert("���Բ��ܳ��� 500 ����!");
	}
	else {
	used.value = message.value.length;
	remain.value = max - used.value;
	}
}
</script>
<style>
<!--
input{border:1px solid #333}
-->
</style>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td> 
      <table width="100%" cellpadding="1" cellspacing="0" class="border">
        <tr class="title"> 
          <td colspan="4">&nbsp;&nbsp;&nbsp;&nbsp;<font color=green><%=PageTitle%></font></td>
        </tr>
        <tr class="tdbg_leftall"> 
          <td colspan="4" align="center" height="10"></td>
        </tr>
        <form name="formwrite" method="post" action="<%=strFileName%>?action=savewrite" onSubmit="return check(formwrite)">
              <input type="hidden" name="reg" value="1">
          <% if WriteType=0 then%>
          <tr class="tdbg_leftall"> 
            <td width="15%" align="center">�� &nbsp;��:</td>
            <td width="35%" align="left"><input type="text" name="GuestName" maxlength="14" size="20" value="<%=WriteName%>"><font color=red>*</font></td>
            <td width="22%" align="left">&nbsp; </td>
            <td width="28%" align="left">&nbsp; </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">��&nbsp;&nbsp;��:</td>
            <td align="left"> 
              <input type="radio" name="GuestSex" value="1" <% if WriteSex=1 then response.write" checked"%> style="BORDER:0px;">
              ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="radio" name="GuestSex" value="0" <% if WriteSex=0 then response.write" checked"%> style="BORDER:0px;">
            Ů </td>
            <td align="left"> &nbsp;&nbsp; 
              <select name="Image" size="1" onChange="changeimage();" >
                <%
				for i=1 to 9
					response.write "<option value='0"&i&"'>0"&i&"</option>"
				next
				for i=10 to 23
					response.write "<option value='"&i&"'>"&i&"</option>"
				next
				%>
              </select>
            </td>
            <td align="left">��</td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">E-mail</td>
            <td align="left"> 
              <input type="text" name="GuestEmail" maxlength="30" size="20" value="<%=WriteEmail%>">
            </td>
            <td rowspan="4" align="left"> 
              <input type="hidden" name="GuestImages" value="01">
			  <img name=showimages src="<%=GuestPath%><%=WriteImages%>.gif" width="80" height="90" border="0" onClick=window.open("guestselect.asp?action=guestimages","face","width=480,height=400,resizable=1,scrollbars=1") title=���ѡ��ͷ�� style="cursor:hand">
		    </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">QQ/MSN:</td>
            <td align="left"> 
              <input type="text" name="GuestOicq" maxlength="15" size="20" value="<%=WriteOicq%>">
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">��ϵ�绰:</td>
            <td align="left"><input type="text" name="GuestIcq" maxlength="15" size="20" value="<%=WriteIcq%>"><font color=red>*</font></td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">��ϵ��ַ:</td>
            <td align="left"><input type="text" name="GuestMsn" maxlength="40" size="20" value="<%=WriteMsn%>"><font color=red>*</font></td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">������ҳ:</td>
			<td colspan="3" align="left"> 
              <input type="text" name="GuestHomepage" maxlength="80" size="37" value="<%=WriteHomepage%>">
			</td>
          </tr>
          <%else%>
          <%end if%>
          <tr class="tdbg_leftall"> 
            <td align="center">��������:</td>
            <td colspan="3" align="left"> 
              <input type="text" name="GuestTitle" size="37" maxlength="28" value="<%=WriteTitle%>">
			  <font color=red>*</font>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">��������:</td>
            <td colspan="3" align="left"> 
              <%
				for i=1 to 20
					response.write "<input type=""radio"" name=""GuestFace"" value="&i&""
					if i=clng(WriteFace) then response.write " checked"
					response.write " style=""BORDER:0px;width:19;"">"
					response.write "<img src="""&GuestPath&"face"&i&".gif"" width=""19"" height=""19"">"& vbcrlf
					if i mod 10 =0 then response.write "<br>"
				next
				%>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">Ubb��ǩ:</td>
            <td colspan="3" align="left"> 
              <% call showubb()%>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td valign="middle" align="center">��������: <br>
            </td>
            <td colspan="3" align="left" valign="top"> 
              <textarea name="GuestContent" cols="59" rows="6" style="border:1px solid #333" onkeydown=gbcount(this.form.GuestContent,this.form.total,this.form.used,this.form.remain); onkeyup=gbcount(this.form.GuestContent,this.form.total,this.form.used,this.form.remain);><%=WriteContent%></textarea>
            </td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td valign="middle" align="center"></td>
            <td colspan="3" align="left" valign="top"> 
				���������
				  <INPUT disabled maxLength=4 name=total size=3 value=500>
				����������<INPUT disabled maxLength=4 name=used size=3 value=0>
				ʣ��������<INPUT disabled maxLength=4 name=remain size=3 value=500>
			</td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td valign="middle" align="center">�Ƿ����أ�</td>
            <td colspan="3" align="left" valign="top"> 
              <input type="radio" name="GuestIsPrivate" value="no"  checked style="BORDER:0px;">
              ���� 
              <input type="radio" name="GuestIsPrivate" value="yes" style="BORDER:0px;">
            ���� &nbsp;&nbsp;<font color=#009900>*</font> ѡ�����غ󣬴�����ֻ�й���Ա�������߲ſ��Կ�����</td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td align="center">�� ֤ �룺</td>
            <td colspan="3" align="left"><input name="CheckCode" size="6" maxlength="4" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onmouseover="this.style.background='#E1F4EE';" onmouseout="this.style.background='#E1F4EE'" onFocus="this.select(); ">
              ����������� <img src=test.asp>
		    <%'response.write session("getcode")%>    <!--img src="inc/checkcode.asp"--></td>
          </tr>
          <tr class="tdbg_leftall"> 
            <td colspan="4" align="center"  height="40"> 
              <input type="hidden" name="saveedit"  value="<%=SaveEdit%>">
              <input type="hidden" name="saveeditid"  value="<%=SaveEditId%>">
              <input type="submit" name="Submit1" value=" �� ��" >
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="reset" name="Submit3" value=" �� �� " >
            </td>
          </tr>
        </form>
		<input type=hidden name=title value=><input type=hidden name=content value=>
		</form>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td  height="15" align="center" valign="top"> 
            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" class="tdbg_left2"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
end sub

'=================================================
'��������ReplyGuest()
'��  �ã��ظ�����
'��  ������
'=================================================
sub ReplyGuest()
	dim ReplyId
	ReplyId=request("guestid")
	if ReplyId="" then
		call Guest_info("<li>��ָ��Ҫ�ظ�������ID��</li>")
		exit sub
	else
		ReplyId=clng(ReplyId)
		sqlGuest="select * from Guest where GuestId=" & ReplyId
	end if
	set rsGuest=server.createobject("adodb.recordset")
	rsGuest.open sqlGuest,conn,1,1
	if rsGuest.bof and rsGuest.eof then
		response.write "<br><li>û���κ�����</li>"
		exit sub
	else
		WriteTitle="Re: "&rsGuest("GuestTitle")
		call ShowGuestList()
	end if
	rsGuest.close
	set rsGuest=nothing
	call WriteGuest()
end sub


'=================================================
'��������EditGuest()
'��  �ã��༭����
'��  ������
'=================================================
sub EditGuest()
	dim EditId
	EditId=request("guestid")
	if EditId="" then
		call Guest_info("<li>��ָ��Ҫ�༭������ID��</li>")
		exit sub
	else
		EditId=clng(EditId)
		sqlGuest="select * from Guest where GuestId=" & EditId
	end if
	set rsGuest=server.createobject("adodb.recordset")
	rsGuest.open sqlGuest,conn,1,1
	if rsGuest.bof and rsGuest.eof then
		response.write "<br><li>�Ҳ�����ָ�������ԣ�</li>"
		exit sub
	end if
	
	if rsGuest("GuestName")=LoginName and rsGuest("GuestIsPassed")=False then
		WriteName=rsGuest("GuestName")
		WriteType=rsGuest("GuestType")
		WriteSex=rsGuest("GuestSex")
		WriteEmail=rsGuest("GuestEmail")
		WriteOicq=rsGuest("GuestOicq")
		WriteIcq=rsGuest("GuestIcq")
		WriteMsn=rsGuest("GuestMsn")
		WriteHomepage=rsGuest("GuestHomepage")
		WriteFace=rsGuest("GuestFace")
		WriteImages=rsGuest("GuestImages")
		WriteTitle=rsGuest("GuestTitle")
		WriteContent=rsGuest("GuestContent")
		SaveEdit=1
		SaveEditId=EditId
		call ShowGuestList()
		call WriteGuest()
	else
		call Guest_info("<li>�û�ֻ���Ա༭�Լ���������ԣ�������δͨ����ˣ�</li>")
	end if    
	rsGuest.close
	set rsGuest=nothing
end sub

'=================================================
'��������DelGuest()
'��  �ã�ɾ������
'��  ������
'=================================================
sub DelGuest()
	dim delid
	delid=trim(Request("guestid"))
	if delid="" then
		call Guest_info("<li>��ָ��Ҫɾ��������ID��</li>")
		exit sub
	end if
	if instr(delid,",")>0 then
		delid=replace(delid," ","")
		sql="Select * from Guest where GuestID in (" & delid & ")"
	else
		delid=clng(delid)
		sql="select * from Guest where GuestID=" & delid
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.bof and rs.eof then
		response.write "<br><li>�Ҳ�����ָ�������ԣ�</li>"
		exit sub
	end if

	if rs("GuestName")<>LoginName or rs("GuestIsPassed")=True then
		call Guest_info("<li>��û��ʹ�ô˹��ܵ�Ȩ�ޣ�</li>")
	else
		do while not rs.eof
			rs.delete
			rs.update
			rs.movenext
		loop
		rs.close
		set rs=nothing
		call Guest_info("<li>ɾ�����Գɹ���</li>")
	end if    
end sub

'=================================================
'��������ShowGuestbutton()
'��  �ã���ʾ���Թ��ܰ�ť
'��  ������
'=================================================
sub ShowGuestButton()
	response.write "<table width=100% border=0 cellpadding=0 cellspacing=3><tr>"
	response.write "<td>"
	if UserHomepage="" or isnull(UserHomepage) then
		response.write "<img src="&GuestPath&"nourl.gif width=45 height=16 alt="&UserGuestName&"û��������ҳ��ַ border=0>"
	else
		response.write "<a href="&UserHomepage&" target=""_blank"">"
		response.write "<img src="&GuestPath&"url.gif width=45 height=16 alt="&UserHomepage&" border=0></a>"
	end if
	'if Usericq="" or isnull(Usericq) then
	'	response.write "<img src="&GuestPath&"notel.gif width=45 height=16 alt="&UserGuestName&"û�����µ绰���� border=0>" & vbcrlf
	'else
	'	response.write "<img src="&GuestPath&"tel.gif width=45 height=16 alt="&Usericq&" border=0 >" & vbcrlf
	'end if
	if UserEmail="" or isnull(UserEmail) then
		response.write "<img src="&GuestPath&"noemail.gif width=45 height=16 alt="&UserGuestName&"û������Email��ַ border=0>"
	else
		response.write "<a href=mailto:"&UserEmail&">"
		response.write "<img src="&GuestPath&"email.gif width=45 height=16 border=0 alt="&UserEmail&"></a>"
	end if
	'response.write "<img src="&GuestPath&"other.gif width=45 height=16 border=0 onMouseOut=toolTip() onMouseOver=""toolTip('&nbsp;QQ/msn��" & UserOicq & "<br>&nbsp;��ϵ��ַ��" & UserMsn & "')"">"
	response.write "</td>"
	response.write "</tr></table>"
end sub
%>