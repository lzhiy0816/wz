<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="../inc/Email_Cls.asp"-->
<%
Head()
Dim Admin_flag
Admin_flag=",20,"
Founderr=False

Dim XmlDom
Dim FilePath
Dim EmailTopic,EmailBody
FilePath = MyDbPath & "data/SendMailLog.config"
FilePath = Server.MapPath(FilePath)

If not Dvbbs.master or instr(","&session("flag")&",",Admin_flag)=0 Then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	Dvbbs_error()
Else
	Call Main()
	Footer()
End If

Sub Main()
%>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
<tr><th colspan="2" height="23">�û��ʼ�֪ͨ</th></tr>
<tr>
<td width="20%" class="forumrow" align="center">
<button Style="width:80;height:50;border: 1px outset ;">ע������</button>
</td>
<td width="80%" class="forumrowHighlight">
	�ٷ����ʼ��б�ֻ�ᱣ������ʮ����¼��
	<br>��ÿ�η����ʼ��벻Ҫ���ù��࣬Ҫ���ݷ��������������;
	<br>���ʼ��б��������͵ļ�¼����δ������Ŀ�������һ��ִ�з��͡�
	<br>�����������ʼ�������ռ�÷�������Դ���뾡���ڷ������ٵ�ʱ���������������
<!-- <br>��
	<br>�� -->
</td>
</tr>
<tr><td colspan="2" class="forumrowHighlight">
<a href="?">ϵͳȺ���ʼ�</a> | <a href="?Act=ShowLog">Ⱥ���ʼ������¼</a>
</td></tr>
</table>
<%
Select Case Request("Act")
	Case "sendemail" : Call SendStep2()
	Case "ShowLog" : Call ShowLog()
	Case "DelSendLog" : Call DelSendLog()
	Case "SendLog"	: Call SendLog()
	Case Else
	Call SendStep1
End Select
End Sub

'ɾ����¼
Sub DelSendLog()
	Dim DelNodes,DelChildNodes
	Set XmlDom = Server.CreateObject("MSXML.DOMDocument")
	If Not XmlDom.load(FilePath) Then
		ErrMsg = "�ʼ��б���Ϊ�գ�����д���ʼ�����ִ�б�����!"
		Dvbbs_Error()
		Exit Sub
	End If
	'Response.Write Request.Form("DelNodes").count
	For Each DelNodes in Request.Form("DelNodes")
		Set DelChildNodes = XmlDom.DocumentElement.selectSingleNode("SendLog[@AddTime='"&DelNodes&"']")
		If Not (DelChildNodes is nothing) Then
			XmlDom.DocumentElement.RemoveChild(DelChildNodes)
		End If
	Next
	XmlDom.save FilePath
	Set XmlDom=Nothing
	Dv_suc("��ѡ�ļ�¼��ɾ��!")
End Sub

'���ݼ�¼�����ʼ�
Sub SendLog()
	Dim SelNodes,SelChildNodes,SendOrders
	SelNodes = Trim(Request.Form("DelNodes"))
	SendOrders = Trim(Request.Form("SendOrders"))
	If SendOrders="" or Not IsNumeric(SendOrders) Then
		ErrMsg = "����дÿ�η����ʼ��ļ�¼��!"
		Dvbbs_Error()
		Exit Sub
	Else
		SendOrders = Clng(SendOrders)
	End If
	Set XmlDom = Server.CreateObject("MSXML.DOMDocument")
	If Not XmlDom.load(FilePath) Then
		ErrMsg = "�ʼ��б���Ϊ�գ�����д���ʼ�����ִ�б�����!"
		Dvbbs_Error()
		Exit Sub
	End If
	Set SelChildNodes = XmlDom.DocumentElement.selectSingleNode("SendLog[@AddTime='"&SelNodes&"']")
	If SelChildNodes is nothing Then
		ErrMsg = "���͵ļ�¼�����ڣ�����д���ʼ�����ִ�б�����!"
		Dvbbs_Error()
		Exit Sub
	End If

	Dim EmailTopic,EmailBody,Total,SearchStr,LastUserID,Remain
	Dim Sql,Rs,i,ii
	Total = SelChildNodes.getAttribute("Total")
	Remain = SelChildNodes.getAttribute("Remain")
	EmailTopic = SelChildNodes.selectSingleNode("EmailTopic").text
	EmailBody = SelChildNodes.selectSingleNode("EmailBody").text
	EmailBody = Replace(EmailBody, CHR(10) & CHR(10), "</P><P> ")
	EmailBody = Replace(EmailBody, CHR(10), "<BR> ")
	SearchStr = SelChildNodes.selectSingleNode("Search").text
	LastUserID = Int(SelChildNodes.getAttribute("LasterUserID"))
	If Remain="0" Then
		ErrMsg = "�Ѿ��������!"
		Dvbbs_Error()
		Exit Sub
	End If
	SQL = "Select Top "&SendOrders&" UserID,UserName,UserEmail From Dv_User where UserID>= " & LastUserID
	If SearchStr<>"" Then
		SQL = SQL &" and "& SearchStr
	End If
	SQL = SQL & " order by UserID "
	SET Rs = Dvbbs.Execute(SQL)
	If Not Rs.eof Then
		SQL=Rs.GetRows(-1)
		Rs.close:Set Rs = Nothing
	Else
		ErrMsg = "�Ѿ��������!"
		Dvbbs_Error()
		Exit Sub
	End If
	%>
	<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
	<tr><td colspan=2 class=forumrow>
	���濪ʼ�����ʼ���Ŀ���û����ܹ�����<%=Total%>�⣬Ŀǰʣ�෢��<%=Remain%>�⣬ÿ�η�������Ϊ<%=SendOrders%>�⡣
	<table width="400" border="0" cellspacing="1" cellpadding="1">
	<tr> 
	<td bgcolor=000000>
	<table width="400" border="0" cellspacing="0" cellpadding="1">
	<tr><td bgcolor=ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
	</td></tr></table>
	<span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
	</table>
	<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
	<tr><td colspan=2 class=forumrow>
	<span id=txt3 name=txt3 style="font-size:9pt">
	</span>
	</td></tr></table>
	<%
	Dim DvEmail
	Set DvEmail = New Dv_SendMail
	DvEmail.SendObject = Cint(Dvbbs.Forum_Setting(2))	'����ѡȡ��� 1=Jmail,2=Cdonts,3=Aspemail
	DvEmail.ServerLoginName = Dvbbs.Forum_info(12)	'�����ʼ���������¼��
	DvEmail.ServerLoginPass = Dvbbs.Forum_info(13)	'��¼����
	DvEmail.SendSMTP = Dvbbs.Forum_info(4)			'SMTP��ַ
	DvEmail.SendFromEmail = Dvbbs.Forum_info(5)		'������Դ��ַ
	DvEmail.SendFromName = Dvbbs.Forum_info(0)		'��������Ϣ
	For i=0 To Ubound(SQL,2)
		If DvEmail.ErrCode = 0 Then
			DvEmail.SendMail SQL(2,i),EmailTopic,EmailBody	'ִ�з����ʼ�
			If Not DvEmail.ErrCode = 0 Then
				ErrMsg = DvEmail.Description
				Dvbbs_Error()
				Exit Sub
			End If
		Else
			ErrMsg = DvEmail.Description
			Dvbbs_Error()
			Exit Sub
		End If
		ii=ii+1
		Response.Write "<script>img2.width=" & Fix((ii/Remain) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""���͸�"&SQL(1,i)&"��"&SQL(2,i)&"�����ʼ���ɣ����ڷ�����һ���û��ʼ���" & FormatNumber(ii/Remain*100,4,-1) & """;" & VbCrLf
		Response.Write "txt3.innerHTML+=""���͸�"&SQL(1,i)&"��"&SQL(2,i)&"�����ʼ����<br>"";"
		Response.Write "</script>"
		Response.Flush
		LastUserID = SQL(0,i)
	Next
	Set DvEmail = Nothing
	Remain = Remain -ii
	If Remain<0 Then Remain = 0
	SelChildNodes.attributes.getNamedItem("Remain").text = Remain
	SelChildNodes.attributes.getNamedItem("LasterUserID").text = LastUserID
	SelChildNodes.attributes.getNamedItem("LastTime").text = now()
	XmlDom.documentElement.appendChild(SelChildNodes)
	XmlDom.save FilePath
	Set XmlDom=Nothing
	If Remain>0 Then
		'�ļ������ͷ�ʽ 2005-10-6 Dv.Yz
		Response.Write "<form method=""POST"" name=""resend"" action=""?Act=SendLog"">"
		Response.Write "<input type=hidden name=""SendOrders"" value=""" & SendOrders & """>"
		Response.Write "<input type=hidden name=""DelNodes"" value=""" & SelNodes & """>"
		Response.Write "&nbsp;&nbsp;<input type=""submit"" value=��������></form>"
	End If
End Sub

'��ʾ�ʼ���¼�б�
Sub ShowLog()
Set XmlDom = Server.CreateObject("MSXML.DOMDocument")
If Not XmlDom.load(FilePath) Then
	ErrMsg = "�ʼ��б���Ϊ�գ�����д���ʼ�����ִ�б�����!"
	Dvbbs_Error()
	Exit Sub
End If
Dim Node,SendLogNode,Childs
Set SendLogNode = XmlDom.DocumentElement.SelectNodes("SendLog")
Childs = SendLogNode.Length	'�б���
If Childs>10 Then
	Dim objRemoveNode,i
	For i=0 To (Childs-11)
	XmlDom.documentElement.removeChild(SendLogNode.item(i))
	Next
	XmlDom.save FilePath
End If
%>
<br>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
<tr><th colspan="9" height="23">�����ʼ��б�</th></tr>
<tr>
<td width="1%" class=bodytitle align=center nowrap>ѡȡ</td>
<td width="20%" class=bodytitle align=center>����</td>
<td width="10%" class=bodytitle align=center nowrap>�ܹ�������Ŀ</td>
<td width="10%" class=bodytitle align=center nowrap>ʣ�෢����Ŀ</td>
<td width="10%" class=bodytitle align=center>������</td>
<td width="10%" class=bodytitle align=center>������IP</td>
<td width="10%" class=bodytitle align=center>���ʱ��</td>
<td width="10%" class=bodytitle align=center>����ʱ��</td>
<td width="10%" class=bodytitle align=center>����</td>
</tr>
<form action="?" method=post name="TheForm">
<tr><td colspan="9" class="forumRowHighlight" height="23">
ÿ�η����ʼ�<INPUT TYPE="text" NAME="SendOrders" value="10" size="4">��
</td></tr>
<%
Dim SearchStr,Topic
i=0
For Each Node in SendLogNode
	'SearchStr = Node.selectSingleNode("Search").text
	Topic = Node.selectSingleNode("EmailTopic").text
	'Node.getAttribute("MasterName")
%>
<tr>
<td class="forumrowHighlight" align=center><INPUT TYPE="checkbox" NAME="DelNodes" value="<%=Node.getAttribute("AddTime")%>"></td>
<td class="forumrow" align=center><%=Topic%></td>
<td class="forumrow"><%=Node.getAttribute("Total")%></td>
<td class="forumrow"><%=Node.getAttribute("Remain")%></td>
<td class="forumrow" align=center><%=Node.getAttribute("MasterName")%></td>
<td class="forumrow"><%=Node.getAttribute("MasterIP")%></td>
<td class="forumrow"><%=Node.getAttribute("AddTime")%></td>
<td class="forumrow"><%=Node.getAttribute("LastTime")%></td>
<td class="forumrow" align=center><input type="submit"  onclick="this.form.Act.value='SendLog';Selchecked(this.form.DelNodes,<%=i%>);" value="����"></td>
</tr>
<%
i=i+1
Next
%>
<tr>
	<td colspan="9" class="forumRowHighlight">
	<input type=hidden name=Act value="DelSendLog">
	<input type=submit name=Submit value="ɾ����¼"  onclick="{if(confirm('ע�⣺��ɾ����ģ�潫���ָܻ���')){this.form.submit();return true;}return false;}">  <input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ȫѡ</td>
</tr>
</form>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
function Selchecked(obj,n){
if (obj[n]){
	obj[n].checked=true;
}else{
	obj.checked=true;
}
}
//-->
</SCRIPT>
<%
Set XmlDom = Nothing
End Sub

'��д�����ʼ���Ϣ
Sub SendStep1()
%>
<br>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
<form METHOD=POST ACTION="?" name="TheForm">
<tr><th colspan="2" height="23">�û��ʼ�֪ͨ</th></tr>
<tr>
<td width="15%" class="forumrowHighlight" align="right">
ѡ���û���
</td>
<td width="85%" class="forumrow">
<INPUT TYPE="text" NAME="UserName" size="40">(����û�������Ӣ�Ķ��š�,���ָ�,ע�����ִ�Сд;)
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
�û����
</td>
<td class="forumrow">
<INPUT TYPE="radio" NAME="UserType" value="0" checked onclick="UType(this.value)">�û�����
<INPUT TYPE="radio" NAME="UserType" value="1" onclick="UType(this.value)">�û���
<INPUT TYPE="radio" NAME="UserType" value="2" onclick="UType(this.value)">�����û�
<div id="ToUserGroup" style="display:none;">
	<br>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
	<tr><td height=20 class="forumrowHighlight">ָ���û���</td></tr>
	<tr><td>
	<%
	'Response.Write "<INPUT TYPE=""checkbox"" NAME=""GetGroupID"" value=""-1"" checked>�����û�"
	Dim Rs
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
	<tr><td height=20 class="forumrowHighlight"><input type="button" value="�򿪸߼�����" NAME="OPENSET" onclick="openset(this,'UpSetting')"></td></tr>
	<tr><td height=20 ID="UpSetting" style="display:NONE" class="forumrowHighlight">
		<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
		<tr><td height=20 colspan="4">������������(����ѡȡ�û��飬�������������������û���Ч)</td></tr>
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
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
�ʼ����⣺
</td>
<td class="forumrow">
<INPUT TYPE="text" NAME="EmailTopic" size="80">
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
�ʼ����ݣ�
</td>
<td class="forumrow">
<TEXTAREA NAME="EmailBody" Style="width:100%;height:250;"></TEXTAREA>
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
</td>
<td class="forumrowHighlight" align="center">
<INPUT TYPE="hidden" name="Act" value="sendemail">
<INPUT TYPE="submit" value="�ύ">&nbsp;&nbsp;&nbsp;<INPUT TYPE="reset" value="����">
</td>
</tr>
</form>
</table>
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
function UType(n){
	var ToUserGroup = document.getElementById("ToUserGroup");
	if (n==0&&TheForm.UserName.disabled==true){
		TheForm.UserName.disabled = false;
		ToUserGroup.style.display = "none";
	}
	else{
		TheForm.UserName.disabled=true;
		if (n==1){
			ToUserGroup.style.display = "";
		}else{
			ToUserGroup.style.display = "none";
		}
	}
}
//-->
</SCRIPT>
<%
End Sub

Sub SendStep2()
	Server.ScriptTimeout=999999
	Dim UserType
	UserType = Request.Form("UserType")
	EmailTopic = Request.Form("EmailTopic")
	EmailBody = Request.Form("EmailBody")
	If EmailTopic="" or EmailBody="" Then
		ErrMsg = "����д�ʼ��ı��������!"
		Dvbbs_Error()
		Exit Sub
	End If
	Select Case UserType
		Case "0" : Call Sendtype_0()	'��ָ���û�
		Case "1" : Call Sendtype_1()	'��ָ���û���
		Case "2" : Call Sendtype_2()	'�������û�
		Case Else
			ErrMsg = "��ѡ���ŵ��û�!"
			Dvbbs_Error()
			Exit Sub
	End Select
	Dv_suc("�Ѿ��ɹ��������¼������б����ڷ����б���ѡȡ����!")
End Sub

'��ָ���û�
Sub Sendtype_0()
	Dim Searchstr
	Dim ToUserName,Rs,Sql,i,ToUserID,FirstUserID
	ToUserName = Trim(Request.Form("UserName"))
	If ToUserName = "" Then
		ErrMsg = "����дĿ���û�����ע�����ִ�Сд��"
		Dvbbs_Error()
		Exit Sub
	End If
	ToUserName = Replace(ToUserName,"'","")
	ToUserName = Split(ToUserName,",")
	If Ubound(ToUserName)>100 Then
		ErrMsg = "����һ�β��ܳ���100λĿ���û���"
		Dvbbs_Error()
		Exit Sub
	End If
	For i=0 To Ubound(ToUserName)
		SQL = "Select UserID From [Dv_user] Where UserName = '"&ToUserName(i)&"' order by userid"
		SET Rs = Dvbbs.Execute(SQL)
		If Not Rs.eof Then
			If i=0 or ToUserID="" Then
				ToUserID = ToUserID & Rs(0)
				FirstUserID = Rs(0)
			Else
				ToUserID = ToUserID &","& Rs(0)
			End If
		End If
	Next
	Rs.Close : Set Rs = Nothing
	Dim Total
	Total = Ubound(Split(ToUserID,","))+1
	If Total = 0 Then
		ErrMsg = "ϵͳ�Ҳ�����ӦĿ���û�����ע�����ִ�Сд��"
		Dvbbs_Error()
		Exit Sub
	Else
		SearchStr = "UserID in ("&ToUserID&")"
		Call CreateXmlLog(Total,SearchStr,FirstUserID)
	End If
End Sub

'��ָ���û��鼰��������
Sub Sendtype_1()
	Dim GetGroupID
	Dim SearchStr,TempValue,DayStr
	GetGroupID = Replace(Request.Form("GetGroupID"),chr(32),"")
	If GetGroupID<>"" and Not Isnumeric(Replace(GetGroupID,",","")) Then
		ErrMsg = "����ȷѡȡ��Ӧ���û��顣"
	Else
		GetGroupID = Dvbbs.Checkstr(GetGroupID)
	End If
	If IsSqlDataBase=1 Then
		DayStr = "d"
	Else
		DayStr = "'d'"
	End If
	If GetGroupID<>"" Then
		If Isnumeric(GetGroupID) Then
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
	If SearchStr="" Then
		ErrMsg = "����д���͵�����ѡ�"
	End If
	If ErrMsg<>"" Then Dvbbs_Error() : Exit Sub
	Dim Rs,Sql,Total,FirstUserID
	Sql = "Select Count(UserID) From Dv_user Where "& SearchStr
	Total = Dvbbs.Execute(Sql)(0)
	If Total>0 Then
		Sql = "Select Top 1 UserID From Dv_user Where "& SearchStr & " order by userid"
		FirstUserID = Dvbbs.Execute(Sql)(0)
		Call CreateXmlLog(Total,SearchStr,FirstUserID)
	Else
		ErrMsg = "����Ŀ���û�Ϊ�գ�����ķ��������ٽ��з��͡�"
		Dvbbs_Error()
		Exit Sub
	End If
End Sub

'�������û�
Sub Sendtype_2()
	Dim SearchStr
	Dim Rs,Sql,Total,FirstUserID
	Sql = "Select Count(UserID) From Dv_user"
	Total = Dvbbs.Execute(Sql)(0)
	If Total>0 Then
		Sql = "Select Top 1 UserID From Dv_user order by userid"
		FirstUserID = Dvbbs.Execute(Sql)(0)
		Call CreateXmlLog(Total,SearchStr,FirstUserID)
	Else
		ErrMsg = "����Ŀ���û�Ϊ�գ�����ķ��������ٽ��з��͡�"
		Dvbbs_Error()
		Exit Sub
	End If
End Sub

'��ӷ��ͼ�¼
Sub CreateXmlLog(SendTotal,Search,LasterUserID)
	Dim node,attributes,createCDATASection,ChildNode
	Set XmlDom = Server.CreateObject("MSXML.DOMDocument")
	If Not XmlDom.load(FilePath) Then
		XmlDom.loadxml "<?xml version=""1.0"" encoding=""gb2312""?><EmailLog/>"
	End If
	Set node=XmlDom.createNode(1,"SendLog","")
	Set attributes=XmlDom.createAttribute("Total")
	attributes.text = SendTotal
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("Remain")
	attributes.text = SendTotal
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("LasterUserID")
	attributes.text = LasterUserID
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("MasterName")
	attributes.text = Dvbbs.Membername
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("MasterUserID")
	attributes.text = Dvbbs.UserID
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("MasterIP")
	attributes.text = Dvbbs.UserTrueIP
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("AddTime")
	attributes.text = Now()
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDom.createAttribute("LastTime")
	attributes.text = Now()
	node.attributes.setNamedItem(attributes)
	Set ChildNode = XmlDom.createNode(1,"Search","")
	Set createCDATASection=XmlDom.createCDATASection(replace(Search,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDom.createNode(1,"EmailTopic","")
	Set createCDATASection=XmlDom.createCDATASection(replace(EmailTopic,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDom.createNode(1,"EmailBody","")
	Set createCDATASection=XmlDom.createCDATASection(replace(EmailBody,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	XmlDom.documentElement.appendChild(node)
	XmlDom.save FilePath
	Set XmlDom = Nothing
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
%>