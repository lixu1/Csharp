<!--#include file = "private.asp"-->

<%

sPosition = sPosition & "�޸��û���������"

Call Header()
Call Content()
Call Footer()


Sub Content()
	Select Case sAction
	Case "MODI"
		Call DoModi()
	Case Else
		Call ShowForm()
	End Select
End Sub


Sub ShowForm()
	%>
	<script language=javascript>
	function checklogin() {
		var obj;
		obj=document.myform.newusr;
		obj.value=BaseTrim(obj.value);
		if (obj.value=="") {
			BaseAlert(obj, "���û�������Ϊ�գ�");
			return false;
		}
		obj=document.myform.newpwd1;
		obj.value=BaseTrim(obj.value);
		if (obj.value=="") {
			BaseAlert(obj, "�����벻��Ϊ�գ�");
			return false;
		}
		if (document.myform.newpwd1.value!=document.myform.newpwd2.value){
			BaseAlert(document.myform.newpwd1, "�������ȷ�����벻��ͬ��");
			return false;
		}
		return true;
	}
	</script>

	<table border=0 cellspacing=1 align=center class=navi>
	<tr><th><%=sPosition%></th></tr>
	</table>

	<br>

	<table border=0 cellspacing=1 align=center class=form>
	<form action='?action=modi' method=post name=myform onsubmit="return checklogin()">
	<tr>
		<th>��������</th>
		<th>������������</th>
		<th>����˵��</th>
	</tr>
	<tr>
		<td width="15%">���û�����</td>
		<td width="55%"><input type=text class=input size=20 name=newusr value="<%=inHTML(Session("eWebEditor_User"))%>"></td>
		<td width="30%"><span class=red>*</span>&nbsp;&nbsp;���û�����<span class=blue><%=outHTML(Session("eWebEditor_User"))%></span></td>
	</tr>
	<tr>
		<td width="15%">�� �� �룺</td>
		<td width="55%"><input type=password class=input size=20 name=newpwd1 maxlength=30></td>
		<td width="30%"><span class=red>*</span></td>
	</tr>
	<tr>
		<td width="15%">ȷ�����룺</td>
		<td width="55%"><input type=password class=input size=20 name=newpwd2 maxlength=30></td>
		<td width="30%"><span class=red>*</span></td>
	</tr>
	<tr><td align=center colspan=3><input type=submit name=bSubmit value="  �ύ  "></a>&nbsp;<input type=reset name=bReset value="  ����  "></td></tr>
	</form>
	</table>


	<%
End Sub

Sub DoModi()

	Dim sNewUsr, sNewPwd1, sNewPwd2
	sNewUsr = Trim(Request("newusr"))
	sNewPwd1 = Trim(Request("newpwd1"))
	sNewPwd2 = Trim(Request("newpwd2"))

	If sNewUsr = "" Then
		GoError "���û�������Ϊ�գ�"
	End If
	If sNewPwd1 = "" then
		GoError "�����벻��Ϊ�գ�"
	End If
	If sNewPwd1 <> sNewPwd2 Then
		GoError "�������ȷ�����벻��ͬ��"
	End If

	sUsername = sNewUsr
	sPassword = sNewPwd1

	Call WriteConfig()

	%>
	<table border=0 cellspacing=1 align=center class=navi>
	<tr><th><%=sPosition%></th></tr>
	</table>

	<br>

	<table border=0 cellspacing=1 align=center class=list>
	<tr>
		<td>��¼�û����������޸ĳɹ���</td>
	</tr>
	</table>
	<%
End Sub

%>