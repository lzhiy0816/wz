<!--#include file="../Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="../inc/md5.asp" -->
<!-- #include file="../inc/chkinput.asp" -->
<%
Dim admin_flag
admin_flag=",2,"
CheckAdmin(admin_flag)
'Alimama����
Dim Forum_Api,Alimama_Api
Dim Action
Dim Alimama
Dim AdsList

ChkForum_api()
Set Alimama = New Alimama_Apicls
Action = Request("act")
Head()
Top_nav()
Page_main()
Footer()
Set Alimama = Nothing
Set Forum_Api = Nothing
Set Alimama_Api = Nothing
Set AdsList = Nothing

Sub Page_main()
	Select Case Action
		Case "reg"
			Reg()
		Case "login"
			Login()
		Case "addads"
			Addads()
		Case "savereg"
			Savereg()
		Case "saveinfo"
			Saveinfo()
		Case "member_bind"
			member_bind()
		Case "info"
			Info()
		Case "saveads"
			SaveAds()
		Case "adslist"
			MyAdsList()
		Case "restore"
			restore()
		Case "editads"
			Editads()
		Case "saveeditads"
			SaveEditads()
		Case "sso_url"
			SSO_Url()
		Case Else
			If Alimama_Api.getAttribute("memberid")="" Then
				Reg()
			Else
				Info()
			End If
	End Select
End Sub

'������
Sub Restore()
	Dvbbs.Execute("update Dv_setup set Forum_apis=''")
	Response.redirect "plus_adsali.asp"
End Sub

Sub ChkForum_api()
	Dim Rs,XmlDoc
	Set Rs = Dvbbs.Execute("Select top 1 Forum_apis From Dv_Setup")
	XmlDoc = Rs(0)
	Rs.Close
	Set Rs = Nothing
	If IsNull(XmlDoc) or XmlDoc = "" Then
		Creat_Forum_Api()
	Else
		Set Forum_Api = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Forum_Api.LoadXml(XmlDoc)
		Set Alimama_Api = Forum_Api.documentElement.selectSingleNode("alimama")
	End If
End Sub

Sub Creat_Forum_Api()
	Set Forum_Api = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Forum_Api.LoadXml("<forum_api/>")
	Set Alimama_Api = Forum_Api.documentElement.appendChild(Forum_Api.createNode(1,"alimama",""))
	Alimama_Api.setAttribute "memberid",""
	Alimama_Api.setAttribute "email",""
	Alimama_Api.setAttribute "password",""
	Alimama_Api.setAttribute "nickname",""
	Alimama_Api.setAttribute "webname",""
	Update_Forum_Api()
End Sub

Sub Update_Forum_Api()
	Dvbbs.Execute("update Dv_setup set Forum_apis='"&Dvbbs.Checkstr(Forum_Api.xml)&"'")
	'Forum_Api.save Server.mappath("../Forum_Api.xml")
End Sub

Sub Top_nav()
%>
<table cellpadding="3" cellspacing="1" border="0" width="100%" align="center">
<tr><td height="25" class="td_title">վ��Ӫ���ƹ�˵��</td></tr>
<%If Alimama_Api.getAttribute("memberid")<>"" Then%>
<tr><td class="td2">
���ѳɹ���Ϊ��������Ļ�Ա���ʺ��ǣ�<font class="font2"><%=Alimama_Api.getAttribute("email")%></font>
</td></tr>
<%End If%>
<tr><td>
<ul>
<li>�����������ֱ�Ӱ���̳�Ĺ��λ����������������Ӫ��ƽ̨��</li>
<li>�������λ�����̣�1��ע���Ϊ���������Ա --> 2���ǼǱ���̳������Ϣ --> 3���������λ��</li>
</ul>
</td></tr>

</table>
<table cellpadding="3" cellspacing="1" border="0" width="100%" align="center">
<tr><td class="td_title">
<%If Alimama_Api.getAttribute("memberid")="" Then%>
<a href="?act=reg" title="��ͨ���������û��ʺ�">�����ʺ�</a>
| <%End If%>
<%If Alimama_Api.getAttribute("webname")="" Then%>
<a href="?act=info">�Ǽ���վ��Ϣ</a>
<%Else%>
<a href="?act=sso_url&react=0">�鿴��վ��Ϣ</a>
<%End If%>
<!-- | <a href="?act=login">վ���½</a>  -->
| <a href="?act=addads">�������λ</a> 
| <a href="?act=adslist">�ҵĹ��λ</a>
<%If Alimama_Api.getAttribute("memberid")<>"" Then%>
| <a href="?act=sso_url&react=1" target="_blank">���׹���</a>
| <a href="?act=sso_url&react=2" target="_blank">Ԥ������</a>
| <a href="?act=sso_url&react=3" target="_blank">��֧����</a>
| <a href="?act=sso_url&react=4" target="_blank">����</a>
| <a href="?act=sso_url&react=5" target="_blank">���λЧ��</a>
<%End If%>
</td></tr>
</table>
<%
End Sub

Sub SSO_Url()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>�������˰���������ʺź��ٽ��з������λ��Ϣ��"
		Dvbbs_error()
		Exit Sub
	End If
	If Alimama_Api.getAttribute("siteid")="" Then
		Errmsg=ErrMsg + "<BR/>������վ��Ϣ��δ�Ǽǣ���ǼǺ��ٽ��з������λ��"
		Dvbbs_error()
		Exit Sub
	End If
	Dim react
	Dim SsoUrl
	Dim Str
	react = Lcase(Request.QueryString("react"))
	Str = Alimama.ApiUrl & "?" & "act=1&signAccount=" & Alimama_Api.getAttribute("email")&"&tourl="&react
	'Response.Write Str
	Response.Redirect Str
End Sub


'��ʾLoading״̬��ʾ
Sub LoadingDiv()
%>
	<div id="loadingDiv">
	<br/>
	<table border="0" cellspacing="1" cellpadding="3"  align="center" width="90%">
	<tr><td class="td2" style="text-align:center;"><img src="../images/loading.gif" border="0" alt=""/> �����ύ������......</td></tr>
	</table>
	<br/>
	</div>
<%
Response.Flush
End Sub

'�ر�Loading״̬��ʾ
Sub CloseLoadingDiv()
%>
<script language="JavaScript">
<!--
if (document.getElementById('loadingDiv')){
	document.getElementById('loadingDiv').style.display = "none";
}
//-->
</script>
<%
End Sub

Sub Info()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>�������˰���������ʺź��ٽ��з�����վ��Ϣ��"
		Dvbbs_error()
		Exit Sub
	End If
	Dim ali_webname,ali_weburl,ali_webdesc,ali_sex,ali_age,occupationSelect,income,ali_weblike
	
	If Alimama_Api.getAttribute("webname")="" Then
		ali_webname = Server.HtmlEncode(Dvbbs.Forum_info(0))
		ali_weburl = Dvbbs.Forum_info(1)
		ali_webdesc = Server.HtmlEncode(Dvbbs.Forum_info(10))
	Else
		ali_webname = Alimama_Api.getAttribute("webname")
		ali_weburl = Alimama_Api.getAttribute("weburl")
		ali_webdesc = Alimama_Api.getAttribute("webdescription")
	End If

%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<tr>
<th colspan="2" style="text-align:center;">
��վ������Ϣ
</th></tr>
<tr><td class="td2 title" colspan="2">����������վ����</td></tr>
<form method="post" action="?act=saveinfo" name="webinfo">
<tr>
<td align="right" width="20%">��վ���ƣ�</td>
<td width="80%"><input type="text" name="ali_webname" size="20" value="<%=ali_webname%>"/>  <font class="font1">��վ�����벻Ҫ����16�����ֻ�32���ַ�</font></td>
</tr>
<tr>
<td align="right">��վ��ַ��</td>
<td><input type="text" name="ali_weburl" size="40" value="<%=ali_weburl%>"/>  <font class="font1">����http://www.dvbbs.net ��վ��ַ�벻Ҫ����255���ַ�</font></td>
</tr>
<tr>
<td align="right">��վ������</td>
<td><textarea name="ali_webdesc" style="width:96%;height:80px;"><%=ali_webdesc%></textarea>  <br/><font class="font1">�����ҵ���վ�ն���IP������PV������PR��4��Alexa����ǰ1���ǹ���רҵ��Ӣ��ѧϰ��վ����ͬ��������ǰ���������ȶ���ע���û�Ⱥ������ÿ�������û�Լ�����ˡ���ץϺ�϶����ҵ���վ���û��Ѿ���3000�ˡ�</font></td>
</tr>
<tr><td class="td2 title" colspan="2">�޸���վ�����Ⱥ</td></tr>

<tr>
<td align="right">�Ա�<font class="font2">*</font> </td>
<td><input type="radio" name="ali_sex" value="1" class="radio"/>�������е�
<input type="radio" name="ali_sex" value="2" class="radio"/>��Ů����һ��
<input type="radio" name="ali_sex" value="3" class="radio"/>������Ů��
</td>
</tr>
<tr>
<td align="right">���䣺<font class="font2">*</font> </td>
<td>
<input type="radio" name="ali_age" value="1" class="radio"/>18������
<input type="radio" name="ali_age" value="2" class="radio"/>18~25��
<input type="radio" name="ali_age" value="3" class="radio"/>25~45��
<input type="radio" name="ali_age" value="4" class="radio"/>45������
</td>
</tr>
<tr>
<td align="right">ְҵ��<font class="font2">*</font> </td>
<td>
<input type="checkbox" name="occupationSelect" value="1" class="checkbox"/>ѧ��
<input type="checkbox" name="occupationSelect" value="2" class="checkbox"/>ְԱ
<input type="checkbox" name="occupationSelect" value="3" class="checkbox"/>�߼�������Ա
<input type="checkbox" name="occupationSelect" value="4" class="checkbox"/>���幤�̻�
<input type="checkbox" name="occupationSelect" value="5" class="checkbox"/>��ҵ��
<input type="checkbox" name="occupationSelect" value="6" class="checkbox"/>����ְҵ��
<input type="checkbox" name="occupationSelect" value="7" class="checkbox"/>��ҵ��
<input type="checkbox" name="selall" value="1" class="checkbox" onclick="boxCheckAll(document.webinfo.occupationSelect,this);"/>����
</td>
</tr>

<tr>
<td align="right">�����룺<font class="font2">*</font> </td>
<td>
<input type="radio" name="income" value="1" class="radio"/>1000����
<input type="radio" name="income" value="2" class="radio"/>1000~3000
<input type="radio" name="income" value="3" class="radio"/>3000~5000
<input type="radio" name="income" value="4" class="radio"/>5000~8000
<input type="radio" name="income" value="5" class="radio"/>8000~10000
<input type="radio" name="income" value="6" class="radio"/>10000����
</td>
</tr>
<tr>
<td align="right">��Ȥ���ã�</td>
<td><input type="text" name="ali_weblike" size="40" value="<%=ali_weblike%>"/>  <font class="font1">��Ȥ�����벻Ҫ����30�����ֻ���60���ַ�</font></td>
</tr>
<%If Alimama_Api.getAttribute("webname")="" Then%>
<tr><td class="td2" colspan="2" align="center"><input type="submit" name="submit" value="ȷ���ύ"/></td></tr>
<%End If%>
</form>
</table>
<script language="JavaScript">
<!--
chkradio(document.webinfo.ali_sex,'<%=Alimama_Api.getAttribute("websex")%>');
chkradio(document.webinfo.ali_age,'<%=Alimama_Api.getAttribute("webage")%>');
chkradio(document.webinfo.income,'<%=Alimama_Api.getAttribute("webincome")%>');
chkcheckbox(document.webinfo.occupationSelect,'<%=Alimama_Api.getAttribute("webjob")%>');
//-->
</script>
<%

End Sub

'����������վ��Ϣ
Sub Saveinfo()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>�������˰���������ʺź��ٽ��з�����վ��Ϣ��"
		Dvbbs_error()
		Exit Sub
	End If
	Dim ali_webname,ali_weburl,ali_webdesc,ali_sex,ali_age,occupationSelect,income,ali_weblike
	ali_webname = Trim(Request.Form("ali_webname"))
	ali_weburl = Trim(Request.Form("ali_weburl"))
	ali_webdesc = Trim(Request.Form("ali_webdesc"))
	ali_sex = Trim(Request.Form("ali_sex"))
	ali_age = Trim(Request.Form("ali_age"))
	occupationSelect = Replace(Trim(Request.Form("occupationSelect"))," ","")
	income = Trim(Request.Form("income"))
	ali_weblike = Trim(Request.Form("ali_weblike"))

	If ali_webname = "" or Len(ali_webname)>32 Then
		Errmsg=ErrMsg + "<BR/>��վ�������벻Ҫ����16�����ֻ�32���ַ���"
		Dvbbs_error()
		Exit Sub
	End If
	If Len(ali_weburl)<3 or Len(ali_weburl)>255 Then
		Errmsg=ErrMsg + "<BR/>��վ�ĵ�ַ�������ַ����ƣ�"
		Dvbbs_error()
		Exit Sub
	End If
	If ali_webdesc = "" Then
		Errmsg=ErrMsg + "<BR/>��վ����������Ϊ�գ�"
		Dvbbs_error()
		Exit Sub
	End If
	If ali_sex = "" or Not Isnumeric(ali_sex) Then
		Errmsg=ErrMsg + "<BR/>��ѡȡ��վ�����Ⱥ�Ա�"
		Dvbbs_error()
		Exit Sub
	End If
	If ali_age = "" or Not Isnumeric(ali_age) Then
		Errmsg=ErrMsg + "<BR/>��ѡȡ��վ�����Ⱥ���䣡"
		Dvbbs_error()
		Exit Sub
	End If
	If occupationSelect = "" or Not IsNumeric(Replace(occupationSelect,",","")) Then
		Errmsg=ErrMsg + "<BR/>��ѡȡ��վ�����Ⱥְҵ��"
		Dvbbs_error()
		Exit Sub
	End If
	If income = "" or Not Isnumeric(income) Then
		Errmsg=ErrMsg + "<BR/>��ѡȡ��վ�����Ⱥ���룡"
		Dvbbs_error()
		Exit Sub
	End If
	If Len(ali_weblike)>60 Then
		Errmsg=ErrMsg + "<BR/>��վ�����Ⱥ����Ȥ���ò��ܳ���60�ַ���"
		Dvbbs_error()
		Exit Sub
	End If


	'-------------------------------------------------------------------------
	'����API�ӿڲ���
	'-------------------------------------------------------------------------
	Dim XmlDom
	Alimama.QueryStr "service","site_add"
	Alimama.QueryStr "memberid",Alimama_Api.getAttribute("memberid")
	Alimama.QueryStr "name",ali_webname
	Alimama.QueryStr "url",ali_weburl
	Alimama.QueryStr "type",1002		'��վ���ͣ�Ĭ��Ϊ1002:������̳
	Alimama.QueryStr "description",ali_webdesc
	Alimama.QueryStr "sex",ali_sex
	Alimama.QueryStr "age",ali_age
	Alimama.QueryStr "income",income
	Alimama.QueryStr "job",occupationSelect
	Alimama.QueryStr "hobby",ali_weblike

	'Response.Write Alimama.Posturl
	'��ʾLoading״̬��ʾ
	LoadingDiv()
	Set XmlDom = Alimama.Post
	If XmlDom is Nothing Then
		CloseLoadingDiv()
		Errmsg=ErrMsg + "<BR/>������󣬻�ȡ�ύ������ֹ�����Ժ����ԣ�"
		Dvbbs_error()
		Exit Sub
	End If
	CloseLoadingDiv()
	'Is_success,Error_code,Data_List
	If Not (Alimama.Is_success is Nothing) Then
		If Alimama.Is_success.text<>"true" Then
			'If Alimama.Error_code.text<"1999" Then
			'	Errmsg=ErrMsg + "<BR/>(Errcode:"&Alimama.Error_code.text&")ϵͳ���󣬻�ȡ�ύ������ֹ�����Ժ����ԣ�"
			'	Dvbbs_error()
			'	Exit Sub
			'End If
			Errmsg=ErrMsg + "<BR/>"&Alimama.Error_desc
			Dvbbs_error()
			Exit Sub
		Else
			If Not (Alimama.Data_List.selectSingleNode("site/siteid") is Nothing) Then
				Alimama_Api.setAttribute "siteid",Alimama.Data_List.selectSingleNode("site/siteid").text
				Alimama_Api.setAttribute "webname",ali_webname
				Alimama_Api.setAttribute "weburl",ali_weburl
				Alimama_Api.setAttribute "webdescription",ali_webdesc
				Alimama_Api.setAttribute "websex",ali_sex
				Alimama_Api.setAttribute "weburl",ali_weburl
				Alimama_Api.setAttribute "webage",ali_age
				Alimama_Api.setAttribute "webincome",income
				Alimama_Api.setAttribute "webjob",occupationSelect
				Alimama_Api.setAttribute "webhobby",ali_weblike
				Update_Forum_Api()
				Dv_suc("��վ��Ϣ�Ǽǳɹ���")
			Else
				Errmsg=ErrMsg + "<BR/>site/siteid�����ڣ�ϵͳδ֪�������Ժ����ԣ�"
				Dvbbs_error()
				Exit Sub
			End If
		End If
	Else
		Errmsg=ErrMsg + "<BR/>�������ݴ��󣬲�������ֹ�����Ժ����ԣ�"
		Dvbbs_error()
		Exit Sub
	End If
End Sub

'�����ʺ�
Sub Reg()
%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<tr><th colspan="2" style="text-align:center;">���������Աע��</th></tr>
<tr><td colspan="2">
<font class="font2">*</font>������Ѿ��ڰ��������������ʺţ�����<a href="?act=login" class="font2">�ʻ���</a>
</td></tr>
<tr><td class="td2 title" colspan="2">���������ʻ���Ϣ</td></tr>
<form method="post" action="?act=savereg">
<tr>
<td align="right">�ǳƣ�</td>
<td><input type="text" name="ali_name" size="20" value="<%=Dvbbs.MemberName%>"/>  <font class="font1">����ǳƽ����ڰ����������̳���������д</font></td>
</tr>
<tr>
<td  align="right">EMAIL��ַ��</td>
<td>
<%
Dim userinfo,UserEmail
If TypeName(Dvbbs.UserSession)="DOMDocument" Then
	Set UserInfo=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo")
	UserEmail = UserInfo.getAttribute("useremail")
Else
	useremail=""
End If
%>
<input type="text" name="ali_email" size="20" value="<%=UserEmail%>"/>
&nbsp;&nbsp;&nbsp;<font class="font1">����ʹ�ô��ʼ���ַ��¼����ȷ��������ʵ��Ч��</font>
</td>
</tr>
<tr>
<td align="right">���룺</td>
<td><input type="text" name="ali_password1" size="20" maxlength="16"/>  <font class="font1">������6-16���ַ���ɣ���ʹ��Ӣ����ĸ�����ֵ��������</font></td>
</tr>
<tr>
<td align="right">������һ�����룺</td>
<td><input type="text" name="ali_password2" size="20" maxlength="16"/>  <font class="font1">��������һ�����������������</font></td>
</tr>
<tr><td class="td2 title" colspan="2">��ϵ��ʽ  (<font class="font2">*����������ϵ��ʽ���ֻ�����Ϊ�������Ϊ��ѡ��</a>)</td></tr>

<tr>
<td  align="right">�ֻ����룺</td>
<td><input type="text" name="ali_mobile" size="20"/>  <font class="font2">*</font><font class="font1">�ֻ������ǰ�������͸��Ͷ�ŵ���Ҫ�������������ϵ����ѡ����������д</font></td>
</tr>
<tr>
<td  align="right">��ϵ�绰��</td>
<td><input type="text" name="ali_phone" size="20"/>  <font class="font1">������������ϵ�绰</font></td>
</tr>
<tr>
<td  align="right">�Ա�������¼����</td>
<td><input type="text" name="ali_wangwang" size="20"/>  <font class="font1">�����������Ա���������Ϊ��֤�͹����ҡ����ҵ�ͨ����ͨ������ʹ�ð���������<a href="http://www.taobao.com/wangwang/index.php" target="_blank">�Ա���</a>��</font></td>
</tr>
<tr><td class="td2" colspan="2" align="center"><input type="submit" name="submit" value="ȷ���ύ"/></td></tr>
</form>
</table>
<%
End Sub

Sub	Savereg()
	If Alimama_Api.getAttribute("memberid")<>"" Then
		Errmsg=ErrMsg + "<BR/>���Ѿ�������ʺ��ˣ�"
		Dvbbs_error()
		Exit Sub
	End If
	Dim Ali_Email,Ali_pass,Ali_name,Ali_mobile
	Dim Ali_phone,Ali_wangwang

	'-------------------------------------------------------------------------
	'��ȡ������
	'-------------------------------------------------------------------------
	Ali_name = Trim(Request.Form("ali_name"))
	Ali_Email = Trim(Request.Form("ali_email"))
	Ali_pass = Trim(Request.Form("ali_password1"))
	Ali_mobile = Trim(Request.Form("ali_mobile"))
	Ali_phone = Trim(Request.Form("ali_phone"))
	Ali_wangwang = Trim(Request.Form("ali_wangwang"))

	'-------------------------------------------------------------------------
	'��֤�ύ���ݿ�ʼ
	'-------------------------------------------------------------------------
	If Ali_name=""or Len(Ali_name)<1 or Len(Ali_name)>32 Then
		Errmsg=ErrMsg + "<BR/>�ǳƲ���Ϊ�ջ򳬹�32���ַ���"
		Dvbbs_error()
		Exit Sub
	End If
	If Len(Ali_Email)>128 or Ali_Email="" Then
		Errmsg=ErrMsg + "<BR/>EMAIL��ַ����Ϊ�ջ򳬹�128���ַ�����������д��"
		Dvbbs_error()
		Exit Sub
	End If
	If IsValidEmail(Ali_Email) = False Then
		Errmsg=ErrMsg + "<BR/>EMAIL��ַ������Ҫ����������д��"
		Dvbbs_error()
		Exit Sub
	End If

	If Ali_pass="" or Len(Ali_pass)<6 or Len(Ali_pass)>16 Then
		Errmsg=ErrMsg + "<BR/>���������6-16���ַ���ɣ���������д��"
		Dvbbs_error()
		Exit Sub
	End If
	If ChkIsEnAndNum(Ali_pass) = False Then
		Errmsg=ErrMsg + "<BR/>��ʹ��Ӣ����ĸ�����ֵ�������룡"
		Dvbbs_error()
		Exit Sub
	End If
	If Ali_pass<>Trim(Request.Form("ali_password1")) Then
		Errmsg=ErrMsg + "<BR/>������������벻������������д��"
		Dvbbs_error()
		Exit Sub
	End If
	If Ali_mobile="" Or Len(Ali_mobile)>50 Then
		Errmsg=ErrMsg + "<BR/>�ֻ�����Ϊ�ջ򳬹�50���ַ���"
		Dvbbs_error()
		Exit Sub
	End If
	If ChkIsPhone(Ali_mobile) = False Then
		Errmsg=ErrMsg + "<BR/>�ֻ����벻����Ҫ������һ�£�"
		Dvbbs_error()
		Exit Sub
	End If
	If Ali_phone<>"" and Len(Ali_phone)>50 Then
		Errmsg=ErrMsg + "<BR/>��ϵ�绰���ܳ���50���ַ�������һ�£�"
		Dvbbs_error()
		Exit Sub
	End If
	
	If Ali_wangwang<>"" and Len(Ali_wangwang)>100 Then
		Errmsg=ErrMsg + "<BR/>�Ա�������¼�����ܳ���100���ַ�������һ�£�"
		Dvbbs_error()
		Exit Sub
	End If
	'��֤�ύ���ݽ���
	'-------------------------------------------------------------------------
	Ali_Email = Lcase(Ali_Email)
	Ali_pass = Md5(Ali_pass,32)

	'-------------------------------------------------------------------------
	'����API�ӿڲ���
	'-------------------------------------------------------------------------
	Dim XmlDom
	Alimama.QueryStr "service","member_register"
	Alimama.QueryStr "logname",Ali_Email
	Alimama.QueryStr "logid",Alimama.FormatNickName(Ali_Email)
	Alimama.QueryStr "passwd",Ali_pass
	Alimama.QueryStr "nickname",Alimama.FormatNickName(Ali_name)
	Alimama.QueryStr "mobile",Ali_mobile
	Alimama.QueryStr "telephone",ali_phone
	Alimama.QueryStr "wangwang",ali_wangwang
	'Response.Write Alimama.Posturl
	'��ʾLoading״̬��ʾ
	LoadingDiv()
	Set XmlDom = Alimama.Post
	If XmlDom is Nothing Then
		CloseLoadingDiv()
		Errmsg=ErrMsg + "<BR/>������󣬻�ȡ�ύ������ֹ�����Ժ����ԣ�"
		Dvbbs_error()
		Exit Sub
	End If
	CloseLoadingDiv()
	'Is_success,Error_code,Data_List
	If Not (Alimama.Is_success is Nothing) Then
		If Alimama.Is_success.text<>"true" Then
			'If Alimama.Error_code.text<"1999" Then
			'	Errmsg=ErrMsg + "<BR/>(Errcode:"&Alimama.Error_code.text&")ϵͳ���󣬻�ȡ�ύ������ֹ�����Ժ����ԣ�"
			'	Dvbbs_error()
			'	Exit Sub
			'End If
			Errmsg=ErrMsg + "<BR/>"&Alimama.Error_desc
			Dvbbs_error()
			Exit Sub
		Else
			Alimama_Api.setAttribute "memberid",Alimama.Data_List.selectSingleNode("member/memberid").text
			Alimama_Api.setAttribute "email",Ali_Email
			Alimama_Api.setAttribute "password",Ali_pass
			Alimama_Api.setAttribute "nickname",Ali_name
			Update_Forum_Api()
			Dv_suc("�����ʺųɹ���")
		End If
	Else
		Errmsg=ErrMsg + "<BR/>�������ݴ��󣬲�������ֹ�����Ժ����ԣ�"
		Dvbbs_error()
		Exit Sub
	End If
	
End Sub

'���û�
Sub Login()
%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<form method="post" action="?act=member_bind">
<tr><th colspan="2" style="text-align:center;">�󶨰��������Ա�ʺ�</th></tr>
<tr><td class="td2 title" colspan="2">��д���ڰ��������ѿ�ͨ���ʻ���Ϣ</td></tr>
<tr>
<td  align="right">EMAIL��ַ��</td>
<td><input type="text" name="ali_email" size="20"/></td>
</tr>
<tr>
<td  align="right">���룺</td>
<td><input type="text" name="ali_password" size="20"/></td>
</tr>
<tr><td class="td2" colspan="2" align="center"><input type="submit" name="submit" value="���ڰ�"/></td></tr>
</form>
</table>
<%
End Sub

'���û�
Sub member_bind()
	Dim Ali_Email,Ali_pass
	Ali_Email = Trim(Request.Form("ali_email"))
	Ali_pass = Trim(Request.Form("ali_password"))
	Ali_Email = Lcase(Ali_Email)
	Ali_pass = Md5(Ali_pass,32)

	'-------------------------------------------------------------------------
	'����API�ӿڲ���
	'-------------------------------------------------------------------------
	Dim XmlDom
	Alimama.QueryStr "service","member_bind"
	Alimama.QueryStr "logname",Ali_Email
	Alimama.QueryStr "passwd",Ali_pass
	Alimama.QueryStr "logid",Alimama.FormatNickName(Ali_Email)
	
	'��ʾLoading״̬��ʾ
	LoadingDiv()
	Set XmlDom = Alimama.Post
	If XmlDom is Nothing Then
		CloseLoadingDiv()
		Errmsg=ErrMsg + "<BR/>������󣬻�ȡ�ύ������ֹ�����Ժ����ԣ�"
		Dvbbs_error()
		Exit Sub
	End If
	CloseLoadingDiv()
	'Is_success,Error_code,Data_List
	If Not (Alimama.Is_success is Nothing) Then
		If Alimama.Is_success.text<>"true" Then
			'If Alimama.Error_code.text<"1999" Then
			'	Errmsg=ErrMsg + "<BR/>(Errcode:"&Alimama.Error_code.text&")ϵͳ���󣬻�ȡ�ύ������ֹ�����Ժ����ԣ�"
			'	Dvbbs_error()
			'	Exit Sub
			'End If
			Errmsg=ErrMsg + "<BR/>"&Alimama.Error_desc
			Dvbbs_error()
			Exit Sub
		Else
			Alimama_Api.setAttribute "memberid",Alimama.Data_List.selectSingleNode("member/memberid").text
			Alimama_Api.setAttribute "email",Ali_Email
			Alimama_Api.setAttribute "password",Ali_pass
			Alimama_Api.setAttribute "nickname",""
			Update_Forum_Api()
			Dv_suc("�ʺŰ󶨳ɹ���")
		End If
	Else
		Errmsg=ErrMsg + "<BR/>�������ݴ��󣬲�������ֹ�����Ժ����ԣ�"
		Dvbbs_error()
		Exit Sub
	End If

End Sub

'���/�༭���λ
Sub Addads()
	Dim Adzoneid,ActStats
	Adzoneid = Request.QueryString("adzoneid")
	If Adzoneid<>"" Then
		Set AdsList = Alimama_Api.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
		If AdsList is Nothing  Then
			Errmsg=ErrMsg + "<BR/>����Ҫ�༭�Ĺ��λ�����ڣ�"
			Dvbbs_error()
			Exit Sub
		End If
		ActStats = "�༭"
	Else
		ActStats = "����"
		Set AdsList = Alimama_Api.appendChild(Forum_Api.createNode(1,"ads",""))
	End If
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>�������˰���������ʺź��ٽ��з������λ��Ϣ��"
		Dvbbs_error()
		Exit Sub
	End If
	If Alimama_Api.getAttribute("siteid")="" Then
		Errmsg=ErrMsg + "<BR/>������վ��Ϣ��δ�Ǽǣ���ǼǺ��ٽ��з������λ��"
		Dvbbs_error()
		Exit Sub
	End If
%>
<style>
#adsize_label span{color:red;}
</style>
<script language="JavaScript">
function chkNum(oIpt){
	var v = oIpt.value;
	var reNum = /[^0-9\.]/;
	if(reNum.test(v)){
	av = v.split('');
	for(var i=0;i<av.length;i++){
	if(reNum.test(av[i])){
	av[i]='';
	}
	}
	oIpt.value = av.join('');
	}
	return true;
}

function _setdemo(v){
	var obj = document.getElementById('ads_setdemo');
	if (obj){
		obj.src = "skins/images/ads_set"+v+".gif";
	}
}
</script>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<form method="post" action="?act=saveads" name="adsform">
<tr><th colspan="2" style="text-align:center;"><%=ActStats%>���λ��Ϣ(<font class="font2">����Ϊ������</font>)</th></tr>
<tr><td class="td2 title" colspan="2">���λ��Ϣ</td></tr>
<tr>
<td width="15%" align="right">���λ���ƣ�</td>
<td width="85%">
<input type="text" name="zonename" size="20" value="<%=AdsList.getAttribute("name")%>"/>(���磺xxx��վ�������λ) 
<br/>
<font class="font1">���λ���ƽ���ʾ����ң��õ�������������ҿ��������λ��ֵ����������д�����ҳ��Ȳ�Ҫ����16������</font>
</td>
</tr>
<tr>
<td  align="right">ʱ���Ʒ����ã�</td>
<td>
<input type="text" name="weekprice" size="5" onKeyup="return chkNum(this);" value="<%=AdsList.getAttribute("weekprice")%>"/> Ԫ/��
</td>
</tr>
<tr>
<td  align="right">�Ƿ���Ҫ��ˣ�</td>
<td>
<input type="radio" name="needAuditingSelect" class="radio" value="1"/>��Ҫ  <input type="radio" name="needAuditingSelect" checked="true" class="radio" value="0"/>����Ҫ
<img style="cursor:help;" alt="����ѡ����Ҫ��˺��н��׸���ʱ��ϵͳ�ᷢ���ʼ�֪ͨ�������ҵĹ���ơ�������ڽ��״�����48Сʱ��δ��ˡ�ϵͳ���Զ�Ϊ�����ͨ����" id="sh_tips" src="skins/images/help.gif" align="absmiddle" />
</td>
</tr>
<tr>
<td  align="right">�����ʽ��</td>
<td>
<input type="radio" name="format" class="radio" value="1"/>���ֹ��  <input type="radio" name="format" checked="true" class="radio"  value="2"/>ͼƬ���
</td>
</tr>
<tr>
<td  align="right">ѡ����λ�ߴ磺</td>
<td>
	<label id="adsize_label"> 
	<input type="radio" name="adsize" value="11" class="radio"/>������<span>760x90</span>&nbsp;
	<input type="radio" name="adsize" value="12" class="radio"/>������<span>468x60</span>&nbsp;
	<input type="radio" name="adsize" value="13" class="radio"/>������<span>250x60</span>&nbsp;
	<input type="radio" name="adsize" value="21" class="radio"/>��ֱ���<span>120x600</span>
	<br/>
	<input type="radio" name="adsize" value="22" class="radio"/>��ֱ���<span>120x240</span>
	<input type="radio" name="adsize" value="31" class="radio"/>�޷����<span>180x250</span>
	<input type="radio" name="adsize" value="32" class="radio"/>�޷����<span>250x300</span>
	<input type="radio" name="adsize" value="33" class="radio"/>�޷����<span>360x190</span>
	</label>
</td>
</tr>

<tr>
<td  align="right">�����ʾλ�ã�</td>
<td>
<div style="float:left;width:40%;">
<input type="checkbox" name="getskinid" value="1" class="checkbox"/>��̳Ĭ�Ϲ��<font class="font1">����������������������ҳ�棩</font>
<br/>
ѡ����ʾ�İ���<font class="font1">�����밴 CTRL �����ж�ѡ��</font>
<select name="getboard" style="width:98%;" size="12" multiple="true">
<%
	Dim node,BoardNode,ii
	Set BoardNode=Application(Dvbbs.CacheName&"_boardlist").cloneNode(True)
	For each node in BoardNode.documentElement.getElementsByTagName("board")
		Response.Write "<option value="""&node.getAttribute("boardid")&""">"
		Select Case node.getAttribute("depth")
			Case 0
				Response.Write "��"
			Case 1
				Response.Write "&nbsp;&nbsp;��"
		End Select
		If node.getAttribute("depth")>1 Then
			For ii=2 To node.getAttribute("depth")
				Response.Write "&nbsp;&nbsp;��"
			Next
			Response.Write "&nbsp;&nbsp;��"
		End If
		Response.Write node.getAttribute("boardtype")
		Response.Write "</option>"&vbNewline
	Next
%>
</select>
<br/>
</div>
<div style="float:left;width:20%;">
ѡ����ҳ������ʾ��λ��
<br/>
ҳ�涥�����λ��<input type="radio" name="adsset" value="1" class="radio" onclick="_setdemo(this.value);"/><br/>
ҳ��ײ����λ��<input type="radio" name="adsset" value="2" class="radio" onclick="_setdemo(this.value);"/><br/>
<!-- ҳ�����ֹ��λ��<input type="radio" name="adsset" value="6" class="radio" onclick="_setdemo(this.value);"/><br/>
����������ֹ��λ��<input type="radio" name="adsset" value="5" class="radio" onclick="_setdemo(this.value);"/><br/> -->
����¥���������λ��<input type="radio" name="adsset" value="7" class="radio" onclick="_setdemo(this.value);"/><br/>
����¥����߹��λ��<input type="radio" name="adsset" value="8" class="radio" onclick="_setdemo(this.value);"/><br/>
����¥���ұ߹��λ��<input type="radio" name="adsset" value="9" class="radio" onclick="_setdemo(this.value);"/><br/>
����¥���ײ����λ��<input type="radio" name="adsset" value="10" class="radio" onclick="_setdemo(this.value);"/><br/>
</div>
<div style="float:left;width:36%;height:220px;">
<img id="ads_setdemo" src="skins/images/ads_set.gif" border="0" alt="���λ��Ԥ��"/>
</div>
</td>
</tr>

<tr>
<td width="15%" align="right">���λ�ؼ��֣�</td>
<td width="85%">
<input type="text" name="keywords" size="20" value="<%=AdsList.getAttribute("keywords")%>"/>
</td>
</tr>

<tr>
<td  align="right">���λ������</td>
<td>
<textarea name="zonedesc" style="width:96%;height:80px;"><%=AdsList.getAttribute("zonedesc")%></textarea>
<br/><font class="font1">���磺���λ��λ�ã���С���Լ��ù��λ������վ�����ж���ҳչ�֣�������Ⱥ��ʲô���ʺ�Ͷ�����棡</font>
</td>
</tr>
<tr><td class="td2" colspan="2" align="center">
<input type="hidden" name="adzoneid" value="<%=AdsList.getAttribute("adzoneid")%>"/>
<input type="submit" name="submit" value="<%=ActStats%>�ҵĹ��λ"/>
</td></tr>
</form>
</table>
<%If Adzoneid<>"" Then%>
<script language="JavaScript">
chkcheckbox(document.adsform.getskinid,'<%=AdsList.getAttribute("homepage")%>');
chkradio(document.adsform.needAuditingSelect,'<%=AdsList.getAttribute("needAuditingSelect")%>');
chkradio(document.adsform.format,'<%=AdsList.getAttribute("format")%>');
chkradio(document.adsform.adsize,'<%=AdsList.getAttribute("adsize")%>');
ChkSelected(document.adsform.getboard,'<%=AdsList.getAttribute("getboard")%>');
chkradio(document.adsform.adsset,'<%=AdsList.getAttribute("adsset")%>');
_setdemo('<%=AdsList.getAttribute("adsset")%>');
</script>
<%
End If
End Sub

'�༭���λ
Sub Editads()
	Dim Adzoneid
	Adzoneid = Request.QueryString("adzoneid")
	If Adzoneid<>"" Then
		Set AdsList = Alimama_Api.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
		If AdsList is Nothing  Then
			Errmsg=ErrMsg + "<BR/>����Ҫ�༭�Ĺ��λ�����ڣ�"
			Dvbbs_error()
			Exit Sub
		End If
	End If
%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<form method="post" action="?act=saveeditads" name="adsform">
<tr><th colspan="2" style="text-align:center;">���λ����</th></tr>
<tr><td class="td2 title" colspan="2">���λ��Ϣ</td></tr>
<tr>
<td  align="right">���λ���ƣ�</td>
<td>
<input type="text" name="zonename" size="20" value="<%=AdsList.getAttribute("name")%>" disabled/> 
</td>
</tr>
<tr>
<td  align="right">�����ʾλ�ã�</td>
<td>
<div style="float:left;width:40%;">
<input type="checkbox" name="getskinid" value="1" class="checkbox"/>��̳Ĭ�Ϲ��<font class="font1">����������������������ҳ�棩</font>
<br/>
ѡ����ʾ�İ���<font class="font1">�����밴 CTRL �����ж�ѡ��</font>
<select name="getboard" size="12" style="width:98%;" multiple="true">
<%
	Dim node,BoardNode,ii
	Set BoardNode=Application(Dvbbs.CacheName&"_boardlist").cloneNode(True)
	For each node in BoardNode.documentElement.getElementsByTagName("board")
		Response.Write "<option value="""&node.getAttribute("boardid")&""">"
		Select Case node.getAttribute("depth")
			Case 0
				Response.Write "��"
			Case 1
				Response.Write "&nbsp;&nbsp;��"
		End Select
		If node.getAttribute("depth")>1 Then
			For ii=2 To node.getAttribute("depth")
				Response.Write "&nbsp;&nbsp;��"
			Next
			Response.Write "&nbsp;&nbsp;��"
		End If
		Response.Write node.getAttribute("boardtype")
		Response.Write "</option>"&vbNewline
	Next
%>
</select>
<br/>
</div>
<div style="float:left;width:20%;">
ѡ����ҳ������ʾ��λ��
<br/>
ҳ�涥�����λ��<input type="radio" name="adsset" value="1" class="radio" onclick="_setdemo(this.value);"/><br/>
ҳ��ײ����λ��<input type="radio" name="adsset" value="2" class="radio" onclick="_setdemo(this.value);"/><br/>
<!-- ҳ�����ֹ��λ��<input type="radio" name="adsset" value="6" class="radio" onclick="_setdemo(this.value);"/><br/>
����������ֹ��λ��<input type="radio" name="adsset" value="5" class="radio" onclick="_setdemo(this.value);"/><br/> -->
����¥���������λ��<input type="radio" name="adsset" value="7" class="radio" onclick="_setdemo(this.value);"/><br/>
����¥����߹��λ��<input type="radio" name="adsset" value="8" class="radio" onclick="_setdemo(this.value);"/><br/>
����¥���ұ߹��λ��<input type="radio" name="adsset" value="9" class="radio" onclick="_setdemo(this.value);"/><br/>
����¥���ײ����λ��<input type="radio" name="adsset" value="10" class="radio" onclick="_setdemo(this.value);"/><br/>
</div>
<div style="float:left;width:36%;height:220px;">
<img id="ads_setdemo" src="skins/images/ads_set.gif" border="0" alt="���λ��Ԥ��"/>
</div>
</td>
</tr>
<tr><td class="td2" colspan="2" align="center">
<input type="hidden" name="adzoneid" value="<%=AdsList.getAttribute("adzoneid")%>"/>
<input type="submit" name="submit" value="ȷ��"/>
</td></tr>
</form>
</table>
<script language="JavaScript">
function _setdemo(v){
	var obj = document.getElementById('ads_setdemo');
	if (obj){
		obj.src = "skins/images/ads_set"+v+".gif";
	}
}
chkcheckbox(document.adsform.getskinid,'<%=AdsList.getAttribute("homepage")%>');
ChkSelected(document.adsform.getboard,'<%=AdsList.getAttribute("getboard")%>');
chkradio(document.adsform.adsset,'<%=AdsList.getAttribute("adsset")%>');
_setdemo('<%=AdsList.getAttribute("adsset")%>');
</script>
<%
End Sub

'������λ����
Sub SaveEditads()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>�������˰���������ʺź��ٽ��з������λ��Ϣ��"
		Dvbbs_error()
		Exit Sub
	End If
	If Alimama_Api.getAttribute("siteid")="" Then
		Errmsg=ErrMsg + "<BR/>������վ��Ϣ��δ�Ǽǣ���ǼǺ��ٽ��з������λ��"
		Dvbbs_error()
		Exit Sub
	End If
	Dim zonename,adsize,getboard,adsset
	Dim Adzoneid,ActStats
	Adzoneid = Trim(Request.Form("adzoneid"))
	getboard = Trim(Request.Form("getboard"))
	adsset = Trim(Request.Form("adsset"))
	If adsset = "" or Not Isnumeric(adsset) Then
		Errmsg=ErrMsg + "<BR/>��ѡȡ�����ҳ������ʾ��λ�ã�"
		Dvbbs_error()
		Exit Sub
	End If

	Set AdsList = Alimama_Api.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
	If AdsList is Nothing  Then
		Errmsg=ErrMsg + "<BR/>����Ҫ�༭�Ĺ��λ�����ڣ�"
		Dvbbs_error()
		Exit Sub
	End If
	AdsList.setAttribute "adsset",adsset
	AdsList.setAttribute "getboard",getboard
	Update_Forum_Api()
	UpdateAdsSeting()
	Dv_suc("���λ���޸ĳɹ���")
End Sub

'���ù��λ������ҳ����ʾ��λ��
Sub UpdateAdsSeting()
	Dim iSetting,i,Forum_ads
	Dim adsset,adsstr
	Dim Sql
	
	'�������ִ�
	'mm_alimama��Աid_��վid_���λid
	adsset = AdsList.getAttribute("adsset")
	Adsstr = "<script type=""text/JavaScript"">" &_ 
	"var alimama_pid=""mm_"&Alimama_Api.getAttribute("memberid")&"_"&Alimama_Api.getAttribute("siteid")&"_"&AdsList.getAttribute("adzoneid")&"""; " &_
	"var alimama_sizecode=""21""; " & _ 
	"var alimama_width="&AdsList.getAttribute("width")&"; " &_ 
	"var alimama_height="&AdsList.getAttribute("height")&"; " &_ 
	"var alimama_type="&AdsList.getAttribute("format")&"; " &_ 
	"</script> " &_ 
	"<script language=""javaScript"" type=""text/javascript"" src=""http://p.alimama.com/inf.js""></script>"

	Adsstr = Replace(Adsstr,"$","")
	If AdsList.getAttribute("homepage")="1" Then
		For i = 0 To 30
			iSetting = Trim(Dvbbs.Forum_ads(i))
			If (i = 2 or i = 7 or i = 13 or i=12 or i=15 or i = 17) and Dvbbs.Forum_ads(i)="" Then iSetting = 0

			If adsset = "1" and i=0 Then	'����
				iSetting = adsstr
			End If
			If adsset = "2" and i=1 Then	'�ײ�
				iSetting = adsstr
			End If
			If adsset = "7" Then			'����¥���������λ
				If i = 18 Then
					iSetting = 1			'��������¥���������λ
				End If
				If i = 19 Then
					iSetting = Adsstr
				End If
			End If

			If adsset = "8" Then			'����¥����߹��λ
				If i = 22 Then
					iSetting = 1			'��������¥����߹��λ
				End If
				If i = 23 Then
					iSetting = Adsstr
				End If
			End If
			If adsset = "9" Then			'����¥���ұ߹��λ
				If i = 22 Then
					iSetting = 2			'��������¥���ұ߹��λ
				End If
				If i = 23 Then
					iSetting = Adsstr
				End If
			End If
			If adsset = "10" Then			'����¥���ײ����λ
				If i = 20 Then
					iSetting = 1			'��������¥���ײ����λ
				End If
				If i = 21 Then
					iSetting = Adsstr
				End If
			End If
			If i = 0 Then
				Forum_ads = iSetting
			Else
				Forum_ads = Forum_ads & "$" & iSetting
			End If
		Next
		Sql = "Update Dv_Setup Set Forum_ads='"&Replace(Forum_ads,"'","''")&"'"
		Dvbbs.Execute(sql)
	End If
	If AdsList.getAttribute("getboard")<>"" Then
		'�����°������ݣ�ֻ��������Ͷ�Ź����������ԭ�����������
		Dim Rs
		Set Rs = Dvbbs.Execute("select Boardid,Board_Ads from Dv_board where boardid in ("&Dvbbs.Checkstr(AdsList.getAttribute("getboard"))&")")
		do while not rs.eof
			Dvbbs.Forum_ads = Split(Rs(1),"$")
			For i = 0 To 30
				iSetting = Trim(Dvbbs.Forum_ads(i))
				If (i = 2 or i = 7 or i = 13 or i=12 or i=15 or i = 17) and Dvbbs.Forum_ads(i)="" Then iSetting = 0

				If adsset = "1" and i=0 Then	'����
					iSetting = adsstr
				End If
				If adsset = "2" and i=1 Then	'�ײ�
					iSetting = adsstr
				End If
				If adsset = "7" Then			'����¥���������λ
					If i = 18 Then
						iSetting = 1			'��������¥���������λ
					End If
					If i = 19 Then
						iSetting = Adsstr
					End If
				End If

				If adsset = "8" Then			'����¥����߹��λ
					If i = 22 Then
						iSetting = 1			'��������¥����߹��λ
					End If
					If i = 23 Then
						iSetting = Adsstr
					End If
				End If
				If adsset = "9" Then			'����¥���ұ߹��λ
					If i = 22 Then
						iSetting = 2			'��������¥���ұ߹��λ
					End If
					If i = 23 Then
						iSetting = Adsstr
					End If
				End If
				If adsset = "10" Then			'����¥���ײ����λ
					If i = 20 Then
						iSetting = 1			'��������¥���ײ����λ
					End If
					If i = 21 Then
						iSetting = Adsstr
					End If
				End If
				If i = 0 Then
					Forum_ads = iSetting
				Else
					Forum_ads = Forum_ads & "$" & iSetting
				End If
			Next
			Sql = "Update Dv_Board Set Board_Ads='"&Replace(Forum_ads,"'","''")&"' Where BoardID ="&Rs(0)
			Dvbbs.Execute(Sql)
		Rs.movenext
		Loop
		Rs.close
		Set Rs = Nothing
	End If
	RestoreBoardCache()
	Dvbbs.loadSetup()
End Sub


'������λ��Ϣ
Sub SaveAds()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>�������˰���������ʺź��ٽ��з������λ��Ϣ��"
		Dvbbs_error()
		Exit Sub
	End If
	If Alimama_Api.getAttribute("siteid")="" Then
		Errmsg=ErrMsg + "<BR/>������վ��Ϣ��δ�Ǽǣ���ǼǺ��ٽ��з������λ��"
		Dvbbs_error()
		Exit Sub
	End If
	Dim zonename,weekprice,needAuditingSelect,format,adsize,getboard,adsset,zonedesc,keywords
	Dim homepage,transtype
	Dim Adzoneid,ActStats
	Adzoneid = Request.Form("adzoneid")
	zonename = Trim(Request.Form("zonename"))
	weekprice = Trim(Request.Form("weekprice"))
	needAuditingSelect = Trim(Request.Form("needAuditingSelect"))
	format = Trim(Request.Form("format"))
	adsize = Trim(Request.Form("adsize"))
	getboard = Trim(Request.Form("getboard"))
	adsset = Trim(Request.Form("adsset"))
	zonedesc  = Trim(Request.Form("zonedesc"))
	keywords = Trim(Request.Form("keywords"))
	'�ύ��Ϣ��֤
	If zonename=""or Len(zonename)<1 or Len(zonename)>32 Then
		Errmsg=ErrMsg + "<BR/>���λ���Ʋ���Ϊ�ջ򳬹�32���ַ���"
		Dvbbs_error()
		Exit Sub
	End If
	If weekprice="" or Not IsNumeric(weekprice) Then
		Errmsg=ErrMsg + "<BR/>ʱ���Ʒ����ñ���Ϊ�����ַ���"
		Dvbbs_error()
		Exit Sub
	End If
	If Ccur(weekprice)<0.05 Then
		Errmsg=ErrMsg + "<BR/>ʱ���Ʒ����ñ�����ڵ���0.05��"
		Dvbbs_error()
		Exit Sub
	End If
	If needAuditingSelect="" or Not IsNumeric(needAuditingSelect) Then
		Errmsg=ErrMsg + "<BR/>����ȷѡ���Ƿ���Ҫ��ˣ�"
		Dvbbs_error()
		Exit Sub
	End If
	If format="" or Not IsNumeric(format) Then
		Errmsg=ErrMsg + "<BR/>����ȷѡ������ʽ��"
		Dvbbs_error()
		Exit Sub
	End If
	If adsize="" or Not IsNumeric(adsize) Then
		Errmsg=ErrMsg + "<BR/>����ȷѡ����λ�ߴ磡"
		Dvbbs_error()
		Exit Sub
	End If
	If zonedesc="" Then
		Errmsg=ErrMsg + "<BR/>�����������Ϊ�գ�"
		Dvbbs_error()
		Exit Sub
	End If
	If adsset = "" or Not Isnumeric(adsset) Then
		Errmsg=ErrMsg + "<BR/>��ѡȡ�����ҳ������ʾ��λ�ã�"
		Dvbbs_error()
		Exit Sub
	End If
	If keywords = "" or Len(keywords)>16 Then
		Errmsg=ErrMsg + "<BR/>����д���λ�ؼ��ֻ�Ҫ����16���ַ���"
		Dvbbs_error()
		Exit Sub
	End If
	If format = "1" and (adsset="3" or adsset="4") Then
		Errmsg=ErrMsg + "<BR/>���ֹ�����ͣ����������ڸ��������¹̶����λ��"
		dvbbs_error()
		Exit Sub
	End If

	'�Ƿ���ʾ����ҳ����
	homepage = Request.Form("getskinid") '����ҳ:1
	'��transtype  �Ʒ����� 1:cpc 2:cpt 3:cpc+cpt  4:cpm
	transtype = 2
	'-------------------------------------------------------------------------
	'���λ�ü��
	'-------------------------------------------------------------------------
	'���λ����ͬһ��վ�������Ʋ����ظ�
	If Adzoneid<>"" Then
		Set AdsList = Alimama_Api.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
		If AdsList is Nothing  Then
			Errmsg=ErrMsg + "<BR/>����Ҫ�༭�Ĺ��λ�����ڣ�"
			Dvbbs_error()
			Exit Sub
		End If
		ActStats = "�༭"
		If Alimama_Api.selectNodes("ads[@name='"&zonename&"' and @adzoneid != "&Adzoneid&" ]").Length>0 Then
			Errmsg=ErrMsg + "<BR/>���λ���Ʋ������ظ���"
			Dvbbs_error()
			Exit Sub
		End If
		AdsList.setAttribute "updatetime",Now()
		Alimama.QueryStr "service","zone_edit"
		Alimama.QueryStr "adzoneid",AdsList.getAttribute("adzoneid")
	Else
		ActStats = "����"
		Set AdsList = Alimama_Api.selectNodes("ads[@name='"&zonename&"']")
		If AdsList.Length>0 Then
			Errmsg=ErrMsg + "<BR/>���λ���Ʋ������ظ���"
			Dvbbs_error()
			Exit Sub
		End If
		'����һ���µĹ�����ݽڵ�
		Set AdsList = Alimama_Api.appendChild(Forum_Api.createNode(1,"ads",""))
		AdsList.setAttribute "createtime",Now()
		AdsList.setAttribute "updatetime",Now()
		Alimama.QueryStr "service","zone_add"
	End If

	'����ѡȡ���ͣ���ȡ���
	Select Case adsize
		Case "11"
			AdsList.setAttribute "width",760
			AdsList.setAttribute "height",90
		Case "12"
			AdsList.setAttribute "width",468
			AdsList.setAttribute "height",60
		Case "13"
			AdsList.setAttribute "width",250
			AdsList.setAttribute "height",60
		Case "21"
			AdsList.setAttribute "width",120
			AdsList.setAttribute "height",600
		Case "22"
			AdsList.setAttribute "width",120
			AdsList.setAttribute "height",240
		Case "31"
			AdsList.setAttribute "width",180
			AdsList.setAttribute "height",250
		Case "32"
			AdsList.setAttribute "width",250
			AdsList.setAttribute "height",300
		Case "33"
			AdsList.setAttribute "width",360
			AdsList.setAttribute "height",190
		Case Else
			AdsList.setAttribute "width",0
			AdsList.setAttribute "height",0
	End Select

	AdsList.setAttribute "name",zonename
	AdsList.setAttribute "weekprice",weekprice
	AdsList.setAttribute "needAuditingSelect",needAuditingSelect
	AdsList.setAttribute "format",format
	AdsList.setAttribute "adsize",adsize
	AdsList.setAttribute "homepage",homepage
	AdsList.setAttribute "transtype",transtype
	AdsList.setAttribute "adsset",adsset
	AdsList.setAttribute "getboard",getboard
	AdsList.setAttribute "zonedesc",zonedesc
	AdsList.setAttribute "keywords",zonedesc
	AdsList.setAttribute "adzonecatids",1099

	'-------------------------------------------------------------------------
	'����API�ӿڲ���
	'-------------------------------------------------------------------------
	Dim XmlDom
	Alimama.QueryStr "memberid",Alimama_Api.getAttribute("memberid")
	Alimama.QueryStr "siteid",Alimama_Api.getAttribute("siteid")
	Alimama.QueryStr "name",zonename
	Alimama.QueryStr "transtype",transtype
	Alimama.QueryStr "weekprice",weekprice
	Alimama.QueryStr "format",format
	Alimama.QueryStr "sizecode",adsize
	Alimama.QueryStr "needauditing",needAuditingSelect
	Alimama.QueryStr "homepage",homepage
	Alimama.QueryStr "adzonecatids",1099			'?���λ��Ŀ
	Alimama.QueryStr "keywords",keywords		'?�ؼ�������
	Alimama.QueryStr "zonedesc",zonedesc

	'��ʾLoading״̬��ʾ
	LoadingDiv()
	Set XmlDom = Alimama.Post
	If XmlDom is Nothing Then
		CloseLoadingDiv()
		Errmsg=ErrMsg + "<BR/>������󣬻�ȡ�ύ������ֹ�����Ժ����ԣ�"
		Dvbbs_error()
		Exit Sub
	End If
	CloseLoadingDiv()
	'Is_success,Error_code,Data_List
	If Not (Alimama.Is_success is Nothing) Then
		If Alimama.Is_success.text<>"true" Then
			'If Alimama.Error_code.text<"1999" Then
			'	Errmsg=ErrMsg + "<BR/>(Errcode:"&Alimama.Error_code.text&")ϵͳ���󣬻�ȡ�ύ������ֹ�����Ժ����ԣ�"
			'	Dvbbs_error()
			'	Exit Sub			
			'End If
			Errmsg=ErrMsg + "<BR/>"&Alimama.Error_desc
			Dvbbs_error()
			Exit Sub
		Else
			If Adzoneid="" Then
				AdsList.setAttribute "adzoneid",Alimama.Data_List.selectSingleNode("adzone/adzoneid").text
			End If
			Update_Forum_Api()
			UpdateAdsSeting()
			Dv_suc(ActStats&"���λ�ɹ���")
		End If
	Else
		Errmsg=ErrMsg + "<BR/>�������ݴ��󣬲�������ֹ�����Ժ����ԣ�"
		Dvbbs_error()
		Exit Sub
	End If
End Sub

'�鿴���λ�б�
Sub MyAdsList()
	Set AdsList = Alimama_Api.selectNodes("ads")
	If AdsList.Length=0 Then
		Errmsg=ErrMsg + "<BR/>��δ�й��λ���ݣ�"
		Dvbbs_error()
		Exit Sub
	End If
	Dim Node
%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<tr><th colspan="7" style="text-align:center;" colspan="4">�ҵĹ��λ��Ϣ</th></tr>
<tr>
<td class="td2 title" >���λ����</td>
<td class="td2 title" width="10%">���λID</td>
<td class="td2 title" width="10%">���λ��ʽ</td>
<td class="td2 title" width="10%">�۸�(Ԫ/��)</td>
<td class="td2 title" width="15%">����ʱ��</td>
<td class="td2 title" width="15%">����ʱ��</td>
<td class="td2 title" >����</td>
</tr>
<%
	For Each Node In AdsList
%>
<tr>
<td><%=Node.getAttribute("name")%></td>
<td><%=Node.getAttribute("adzoneid")%></td>
<td>cpt</td>
<td><%=Node.getAttribute("weekprice")%></td>
<td>
<%
If IsDate(Node.getAttribute("createtime")) Then
Response.Write Formatdatetime(Node.getAttribute("createtime"),0)
Else
Response.Write Node.getAttribute("createtime")
End If
%></td>
<td>
<%
If IsDate(Node.getAttribute("updatetime")) Then
Response.Write Formatdatetime(Node.getAttribute("updatetime"),0)
Else
Response.Write Node.getAttribute("updatetime")
End If
%></td>
<td><a href="?act=editads&adzoneid=<%=Node.getAttribute("adzoneid")%>">�༭</a></td>
</tr>
<%
	Next
%>
</table>
<%
Set AdsList = Nothing
End Sub

'���°����滺������
Sub RestoreBoardCache()
	Dim Board,node
	Dvbbs.LoadBoardList()
	For Each node in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style/@id")
		Application.Contents.Remove(Dvbbs.CacheName & "_showtextads_"&node.text)
		For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
			Dvbbs.LoadBoardData board.text
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text)
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text&"_-time")
		Next
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text)
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text&"_-time")
	Next
End Sub

'-------------------------------------------------------------------------
'��������API�ӿ���
'-------------------------------------------------------------------------

Class Alimama_Apicls
	Public Return_url	'��ѡ����
	Public ApiUrl,ApiKey,Service,Partner,Sign,Sign_type	'��ѡ��������
	Private Query_Str
	Public Is_success,Error_code,Data_List,Error_desc

	Private Sub Class_Initialize()
		'ApiUrl = "http://192.168.1.103:89/aliads_api.asp"
		ApiUrl = "http://server.dvbbs.net/ads_api.asp"
		Sign_type = "MD5"
		Query_Str = ""
		'QueryStr "sign_type",Sign_type	'���ܷ�ʽ
		'QueryStr "return_url",Dvbbs.Get_ScriptNameUrl	'����url
		Set Is_success = Nothing
		Set Error_code = Nothing
		Set Data_List = Nothing
		Error_desc = ""
	End Sub

	Private Sub Class_Terminate()
	End Sub
	
	'nickname ��ʽ��
	Public Function FormatNickName(Str)
		FormatNickName = Str & "_cndw"
	End Function

	'����URL������������������ֵ��
	Public Function QueryStr(qName,qStr)
		If qStr="" or IsNull(qStr) Then Exit Function
		If Query_Str<>"" Then Query_Str = Query_Str & "&"
		Query_Str = Query_Str & qName & "=" & qStr
	End Function

	'��ȡ����URL��ַ
	Public Property Get PostUrl()
		Dim Str
		Str = SortByEn(Query_Str,"&")
		'Sign = Md5(Str & ApiKey,32)
		'Str = Str & "&sign=" & Sign
		'Str = Str & "&sign_type=" & Sign_type
		PostUrl = ApiUrl & "?" & SortByEn(Str,"&")
		'Response.Write PostUrl
		'Response.Write "<br>"
	End Property

	'��URL�Ĳ�������ĸ����
	Public Function SortByEn(Str,S)
		Dim Sort,QStr
		QStr = Mid(Str,InStrRev(Str, "?")+1)
		Sort = JsSortByEn(QStr,S)
		SortByEn = Left(Str,Instr(Str,"?")) & Sort
	End Function

	Public Function Post()
		Dim XmlHttp,XmlDom
		Set XmlHttp = Server.CreateObject("Microsoft.XMLHTTP")
		XmlHttp.Open "post",PostUrl,false
		XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		'XmlHttp.SetRequestHeader "content-type", "text/xml"
		XmlHttp.send()

		Set XmlDom = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		If XmlDom.Load(xmlhttp.responseXML) Then
			Set Post = XmlDom
			Set Is_success = XmlDom.documentElement.selectSingleNode("is_success")
			
			Set Error_code = XmlDom.documentElement.selectSingleNode("error_code")
			If Not (Error_code is Nothing) Then
				Error_desc = "(Errcode:" & Error_code.text & "):" & Error_code.getAttribute("desc")
			End If
			Set Data_List = XmlDom.documentElement.selectSingleNode("data/list")
		Else
			'Response.Write PostUrl
			Response.Write xmlhttp.responsetext
			Set Post = Nothing
		End If
		Set XmlHttp = Nothing
		Set XmlDom = Nothing
	End Function

End Class

'��������
Function VBSort(ary)
	Dim KeepChecking,I,FirstValue,SecondValue
	 KeepChecking = TRUE 
	Do Until KeepChecking = FALSE 
	 KeepChecking = FALSE 
	 For I = 0 to UBound(ary) 
	  If I = UBound(ary) Then Exit For 
	   If ary(I) > ary(I+1) Then 
		FirstValue = ary(I) 
		SecondValue = ary(I+1) 
		ary(I) = SecondValue 
		ary(I+1) = FirstValue 
		KeepChecking = TRUE 
	   End If 
	 Next 
	Loop 
	 VBSort = ary 
End Function

'��ȡ��ǰʱ��
function get_nowtime()
	dim year_str , month_str , day_str , hour_str , minute_str , second_str
	year_str=cstr(year(now))
	month_str=cstr(month(now))
	day_str=cstr(day(now))
	hour_str=cstr(hour(now))
	minute_str=cstr(minute(now))
	second_str=cstr(second(now))
	if len(hour_str)=1 then hour_str="0"+hour_str
	if len(minute_str)=1 then minute_str="0"+minute_str
	if len(second_str)=1 then second_str="0"+second_str
	get_nowtime=year_str+"-"+month_str+"-"+day_str+" "+hour_str+":"+minute_str+":"+second_str 
end function


%>
<script type="text/javascript" runat="server" language="javascript">
//�ַ�������ĸ����,SΪ�ָ���
function JsSortByEn(Str,S){
   var a, l;
   a = Str.split(S);
   l = a.sort();                   // �������顣
   return(l.join(S));
}
</script>