<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
	Head()
	dim admin_flag
	admin_flag=",18,"
	If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
		dvbbs_error()
	Else
		Main_head()
		Select Case Request("action")
		Case "SendMoney"
			SendMoney
		Case Else
			SendForm
		End Select
		If ErrMsg<>"" Then Dvbbs_Error
		If founderr then call dvbbs_error()
		footer()
	End If

'����˵����ע������
Sub Main_head

End Sub

'�������
Sub SendForm
Dim Rs
%>
<form METHOD=POST ACTION="?action=SendMoney">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><th height=20 colspan="4">������������</th></tr>
<tr>
<td class="forumRow" align=right width="15%"><U>���ͽ��</U>��</td>
<td class="forumRow" width="40%">
<INPUT TYPE="text" NAME="SendMoney" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" NAME="SendMoneyType" checked value="0">���� <INPUT TYPE="radio" NAME="SendMoneyType" value="1">���� <INPUT TYPE="radio" NAME="SendMoneyType" value="2">����
</td>
<td class="forumRow" width="10%"><INPUT TYPE="checkbox" NAME="SelectType" value="SendMoney">ѡȡ</td>
<td class="forumrowHighlight" width="*" rowspan="6" valign=top>
<li>����ȷ��д�����ֵ��
<li>ѡȡ�������²�����Ч��
<li>��ѡȡ���£���Ŀ���û�������ݽ�����Ϊ�����ã�
</td>
</tr>
<tr>
<td class="forumRow" align=right><U>���͵�ȯ</U>��</td>
<td class="forumRow"><INPUT TYPE="text" NAME="SendTicket" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" NAME="SendTicketType" checked value="0">���� <INPUT TYPE="radio" NAME="SendTicketType" value="1">���� <INPUT TYPE="radio" NAME="SendTicketType" value="2">����
</td>
<td class="forumRow"><INPUT TYPE="checkbox" NAME="SelectType" value="SendTicket">ѡȡ</td>
</tr>
<tr>
<td class="forumRow" align=right><U>���;���</U>��</td>
<td class="forumRow"><INPUT TYPE="text" NAME="SendUserEP" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" NAME="SendUserEPType" checked value="0">���� <INPUT TYPE="radio" NAME="SendUserEPType" value="1">���� <INPUT TYPE="radio" NAME="SendUserEPType" value="2">����
</td>
<td class="forumRow"><INPUT TYPE="checkbox" NAME="SelectType" value="SendUserEP">ѡȡ</td>
</tr>
<tr>
<td class="forumRow" align=right><U>��������</U>��</td>
<td class="forumRow"><INPUT TYPE="text" NAME="SendUserCP" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" NAME="SendUserCPType" checked value="0">���� <INPUT TYPE="radio" NAME="SendUserCPType" value="1">���� <INPUT TYPE="radio" NAME="SendUserCPType" value="2">����
</td>
<td class="forumRow"><INPUT TYPE="checkbox" NAME="SelectType" value="SendUserCP">ѡȡ</td>
</tr>
<tr>
<td class="forumRow" align=right><U>���ͽ�Ǯ</U>��</td>
<td class="forumRow"><INPUT TYPE="text" NAME="SendUserWealth" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" NAME="SendUserWealthType" checked value="0">���� <INPUT TYPE="radio" NAME="SendUserWealthType" value="1">���� <INPUT TYPE="radio" NAME="SendUserWealthType" value="2">����
</td>
<td class="forumRow"><INPUT TYPE="checkbox" NAME="SelectType" value="SendUserWealth">ѡȡ</td>
</tr>
<tr>
<td class="forumRow" align=right><U>��������</U>��</td>
<td class="forumRow"><INPUT TYPE="text" NAME="SendUserPower" size=10 onkeyup="CheckNumer(this.value,this,'')">
<INPUT TYPE="radio" NAME="SendUserPowerType" checked value="0">���� <INPUT TYPE="radio" NAME="SendUserPowerType" value="1">���� <INPUT TYPE="radio" NAME="SendUserPowerType" value="2">����
</td>
<td class="forumRow"><INPUT TYPE="checkbox" NAME="SelectType" value="SendUserPower">ѡȡ</td>
</tr>
<tr><th height=20 colspan="4">��������Ŀ��</th></tr>
<tr><td class="forumrowHighlight" height=20 colspan="4">
<INPUT TYPE="radio" NAME="Sendtype" value="0" onclick="formstep(0)">��ָ���û�
<INPUT TYPE="radio" NAME="Sendtype" value="1" onclick="formstep(1)">��ָ���û���
<INPUT TYPE="radio" NAME="Sendtype" value="2" onclick="formstep(2)">�������û�
</td></tr>
</table>
<div id="ToUser" style="display:none;">
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr><th height=20 colspan="2">ָ���û�</th></tr>
	<tr><td height=20 colspan="2">�û�����Ӣ�Ķ��š�,���ָ���Ϊ��ʡ��Դ��ÿ�θ�������10λ�û���ע�����ִ�Сд��</td></tr>
	<td class="forumRow"><u>�û�����</u>��</td>
	<td class="forumRow"><INPUT TYPE="text" NAME="ToUserName" size="80"></td>
	</table>
</div>
<div id="ToUserGroup" style="display:none;">
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
	<tr><th height=20>ָ���û���</th></tr>
	<tr><td>
	<li>��ѡȡָ�����µ��û���<LI>��ֻ��ĳ�����û�����£��벻Ҫѡȡ�����û���<br>
	<%
	Set Rs=DvBBS.Execute("Select UserGroupID,Title,UserTitle,parentgid From Dv_UserGroups where parentgid>0  Order By parentgid,UserGroupID")
	Do while not Rs.eof
		Response.Write "&nbsp;&nbsp;<INPUT TYPE=""checkbox"" NAME=""GetGroupID"" value="""&Rs(0)&""">"
		Response.Write Rs(2)
	Rs.movenext
	Loop
	Rs.close
	Set Rs=Nothing
	%>
	</td></tr>
	<tr><td height=20 class="forumrowHighlight" ><input type="button" value="�򿪸߼�����" NAME="OPENSET" onclick="openset(this,'UpSetting')"></td></tr>
	<tr><td height=20 ID="UpSetting" style="display:NONE" class="forumrowHighlight">
		<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
		<tr><th height=20 colspan="4">������������</th></tr>
		<tr>
		<td class="forumRow" width="15%">����½ʱ�䣺</td>
		<td class="forumRow" width="35%">
		<input type="text" name="LoginTime" onkeyup="CheckNumer(this.value,this,'')" size=6>�� &nbsp;<INPUT TYPE="radio" NAME="LoginTimeType" checked value="0">���� <INPUT TYPE="radio" NAME="LoginTimeType" value="1">����
		</td>
		<td class="forumRow" width="15%">ע��ʱ�䣺</td>
		<td class="forumRow" width="35%">
		<input type="text" name="RegTime" onkeyup="CheckNumer(this.value,this,'')" size=6>�� &nbsp;<INPUT TYPE="radio" NAME="RegTimeType" checked value="0">���� <INPUT TYPE="radio" NAME="RegTimeType" value="1">����
		</td>
		</tr>
		<tr>
		<td class="forumRow">��½������</td>
		<td class="forumRow"><input type="text" name="Logins" size=6 onkeyup="CheckNumer(this.value,this,'')">�� &nbsp;<INPUT TYPE="radio" NAME="LoginsType" checked value="0">���� <INPUT TYPE="radio" NAME="LoginsType" value="1">����
		</td>
		<td class="forumRow">�������£�</td>
		<td class="forumRow"><input type="text" name="UserPost" size=6 onkeyup="CheckNumer(this.value,this,'')">ƪ &nbsp;<INPUT TYPE="radio" NAME="UserPostType" checked value="0">���� <INPUT TYPE="radio" NAME="UserPostType" value="1">����</td>
		</tr>
		<tr>
		<td class="forumRow">�������£�</td>
		<td class="forumRow"><input type="text" name="UserTopic" size=6 onkeyup="CheckNumer(this.value,this,'')">ƪ &nbsp;<INPUT TYPE="radio" NAME="UserTopicType" checked value="0">���� <INPUT TYPE="radio" NAME="UserTopicType" value="1">����</td>
		<td class="forumRow">�������£�</td>
		<td class="forumRow"><input type="text" name="UserBest" size=6 onkeyup="CheckNumer(this.value,this,'')">ƪ &nbsp;<INPUT TYPE="radio" NAME="UserBestType" checked value="0">���� <INPUT TYPE="radio" NAME="UserBestType" value="1">����
		</td>
		</tr>
		</table>
	</td></tr>
	</table>
</div>
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
	<tr><td height=20 align=center><input type="submit" value="ִ�и���"></td></tr>
</table>
<form>
<SCRIPT LANGUAGE="JavaScript">
<!--
function openset(v,s){
	if (v.value=='�򿪸߼�����'){
		document.getElementById(s).style.display = "";
		v.value="�رո߼�����";
	}
	else{
		v.value="�򿪸߼�����";
		document.getElementById(s).style.display = "none";
	}
}
//��֤������ֵ n:number ����ֵ | v:value object �������� | n_max ���ֵ
function CheckNumer(n,v,n_max)
{
	if (isNaN(n)){
		v.value = "";
		alert("����д��ȷ����ֵ��");
	}
	else{
		n = parseInt(n);
		if (!isNaN(n_max)){
			n_max = parseInt(n_max);
			if (n>n_max){v.value = "";alert("������ֵ���ܸ��ڣ�"+n_max);}
		}
	}
}

function formstep(OpenID){
	var ToUser = document.getElementById("ToUser");
	var ToUserGroup = document.getElementById("ToUserGroup");
	if (OpenID==0){
	ToUser.style.display = "";
	ToUserGroup.style.display = "none";
	}
	else if (OpenID==1){
	ToUser.style.display = "none";
	ToUserGroup.style.display = "";
	}
	else{
	ToUser.style.display = "none";
	ToUserGroup.style.display = "none";
	}
}
//-->
</SCRIPT>
<%
End Sub

'�����������
Sub SendMoney
	Dim SelectType,UPString,TempData
	SelectType = Replace(Request.Form("SelectType"),chr(32),"")
	If SelectType="" Then
		ErrMsg = "��ѡȡ����������!"
		Exit Sub
	End If
	SelectType = ","&SelectType&","
	UPString = ""
	'���½��
	If Instr(SelectType,"SendMoney") Then
		UPString = GetUPString(Request.Form("SendMoney"),UPString,Request.Form("SendMoneyType"),"UserMoney")
	End If
	'���µ�ȯ
	If Instr(SelectType,"SendTicket") Then
		UPString = GetUPString(Request.Form("SendTicket"),UPString,Request.Form("SendTicketType"),"UserTicket")
	End If
	'���¾���
	If Instr(SelectType,"SendUserEP") Then
		UPString = GetUPString(Request.Form("SendUserEP"),UPString,Request.Form("SendUserEPType"),"UserEP")
	End If
	'��������
	If Instr(SelectType,"SendUserCP") Then
		UPString = GetUPString(Request.Form("SendUserCP"),UPString,Request.Form("SendUserCPType"),"UserCP")
	End If
	'���½�Ǯ
	If Instr(SelectType,"SendUserWealth") Then
		UPString = GetUPString(Request.Form("SendUserWealth"),UPString,Request.Form("SendUserWealthType"),"UserWealth")
	End If
	'��������
	If Instr(SelectType,"SendUserPower") Then
		UPString = GetUPString(Request.Form("SendUserPower"),UPString,Request.Form("SendUserPowerType"),"UserPower")
	End If
	'Response.Write UPString
	Select Case Request.Form("Sendtype")
		Case "0" : Call Sendtype_0(UPString)	'��ָ���û�
		Case "1" : Call Sendtype_1(UPString)	'��ָ���û���
		Case "2" : Call Sendtype_2(UPString)	'�������û�
		Case Else
			ErrMsg = "��ѡȡ��������Ŀ��!"
			Exit Sub
	End Select
End Sub

'��ָ���û�
Sub Sendtype_0(Str)
	Dim ToUserName,Rs,Sql,i,ToUserID
	ToUserName = Trim(Request.Form("ToUserName"))
	If ToUserName = "" Then ErrMsg = "����дĿ���û�����ע�����ִ�Сд��" : Exit Sub
	ToUserName = Replace(ToUserName,"'","")
	ToUserName = Split(ToUserName,",")
	If Ubound(ToUserName)>10 Then ErrMsg = "����һ�β��ܳ���10λĿ���û���" : Exit Sub
	For i=0 To Ubound(ToUserName)
		SQL = "Select UserID From [Dv_user] Where UserName = '"&ToUserName(i)&"'"
		SET Rs = Dvbbs.Execute(SQL)
		If Not Rs.eof Then
			If i=0 or ToUserID="" Then
				ToUserID = ToUserID & Rs(0)
			Else
				ToUserID = ToUserID &","& Rs(0)
			End If
		Else
			ErrMsg = "Ŀ���û������ڣ�ע�����ִ�Сд��" : Exit Sub
		End If
	Next
	Rs.Close : Set Rs = Nothing
	If ToUserID<>"" Then
		SQL = "Update [Dv_user] Set "&Dvbbs.Checkstr(Str)&" where UserID in ("&ToUserID&") "
		Dvbbs.Execute(SQL)
		Dv_suc("��λ"&Ubound(ToUserName)+1&"Ŀ���Ա���³ɹ�!")
	Else
		ErrMsg = "Ŀ���û������ڣ�ע�����ִ�Сд��" : Exit Sub
	End If
End Sub

'��ָ���û���
Sub Sendtype_1(Str)
	Dim GetGroupID
	Dim SearchStr,TempValue,DayStr
	GetGroupID = Replace(Request.Form("GetGroupID"),chr(32),"")
	If GetGroupID="" or Not Isnumeric(Replace(GetGroupID,",","")) Then
		ErrMsg = "����ȷѡȡ��Ӧ���û��顣" : Exit Sub
	Else
		GetGroupID = Dvbbs.Checkstr(GetGroupID)
	End If
	If IsSqlDataBase=1 Then
		DayStr = "d"
	Else
		DayStr = "'d'"
	End If
	If Instr(GetGroupID,"-1") Then
		SearchStr = ""
	Else
		If Instr(GetGroupID,",")=0 Then
			SearchStr = "UserGroupID = "&GetGroupID
		Else
			SearchStr = "UserGroupID in ("&GetGroupID&")"
		End If
	End If
	'��½����
	TempValue = Request.Form("Logins")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginsType"),"UserLogins")
	End If
	'��������
	TempValue = Request.Form("UserPost")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserPostType"),"UserPost")
	End If
	'��������
	TempValue = Request.Form("UserTopic")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserTopicType"),"UserTopic")
	End If
	'��������
	TempValue = Request.Form("UserBest")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserBestType"),"UserIsBest")
	End If
	'����½ʱ��
	TempValue = Request.Form("LoginTime")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginTimeType"),"Datediff("&DayStr&",Lastlogin,"&SqlNowString&")")
	End If
	'ע��ʱ��
	TempValue = Request.Form("RegTime")
	If TempValue<>"" and IsNumeric(TempValue) Then
		SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("RegTimeType"),"Datediff("&DayStr&",JoinDate,"&SqlNowString&")")
	End If

	Dim SQL
	SQL = "Update [Dv_user] Set "&Dvbbs.Checkstr(Str)&" Where "&SearchStr
	Dvbbs.Execute(SQL)
	Dv_suc("Ŀ���Ա���³ɹ�!")
End Sub

'�������û�
Sub Sendtype_2(Str)
	SQL = "Update [Dv_user] Set "& Dvbbs.Checkstr(Str)
	Dvbbs.Execute(SQL)
	Dv_suc("���л�Ա���³ɹ�!")
End Sub

Function GetSearchString(Get_Value,Get_SearchStr,UpType,UpColumn)
	Get_Value = Clng(Get_Value)
	If Get_SearchStr<>"" Then Get_SearchStr = Get_SearchStr & " and " 
	If UpType="1" Then
		Get_SearchStr = Get_SearchStr & UpColumn &" <= "&Get_Value
	Else
		Get_SearchStr = Get_SearchStr & UpColumn &" >= "&Get_Value
	End If
	GetSearchString = Get_SearchStr
End Function

Function GetUPString(TempData,UPString,UpType,UpColumn)
	If TempData<>"" and IsNumeric(TempData) Then
			If UPString<>"" Then UPString = UPString & ","
			Select Case UpType
				Case "2" : UPString = UPString &" "&UpColumn&" = "&cCur(TempData)
				Case "1" : UPString = UPString &" "&UpColumn&" = "&UpColumn&"-"&cCur(TempData)
				Case Else : UPString = UPString &" "&UpColumn&" = "&UpColumn&"+"&cCur(TempData)
			End Select
			GetUPString = UPString
	End If
End Function
%>