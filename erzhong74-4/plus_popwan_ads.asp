<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="Plus_popwan/cls_setup.asp"-->
<%
	Dim Action

	Dim Errmsg
	Dim NewsConfigFile
	Dim popwan_ads,Forum_Api,AdsList
	
	Dvbbs.LoadTemplates("")
	Dvbbs.Stats = "Ͷ�Ź���ƹ�"
	Dvbbs.Nav()
	Dvbbs.Head_var 0,0,Plus_Popwan.Program,"plus_popwan_ads.asp"
	Dvbbs.ActiveOnline()
	action = Request("action")
	Page_main()

	If action<>"frameon" Then
		Dvbbs.Footer
	End If
	Dvbbs.PageEnd()

'ҳ���Ҳ����ݲ���
Sub Page_Center()
	If Not (Dvbbs.master Or Dvbbs.GroupSetting(70)="1") Then
		Dvbbs.AddErrcode(28)
		Dvbbs.ShowErr()
	End If

	Dim act
	Act = Request("act")

	NewsConfigFile = Plus_Popwan.FilePath(Plus_Popwan.CachePath&"Temp_Plus_popwan_ads.config")
	ChkForum_api()

	Page_main1()

	Select Case Act
		Case "addads"	'�������λ
			Addads()
		Case "saveads"	'������λ
			Call saveads()
		Case "adslist"	'�ҵĹ��λ
			Call MyAdsList()
		Case "editads"  '�༭���λ
			Call Editads()
		Case "saveeditads"  '����༭���λ
			Call SaveEditads()
		Case "restore"	'������
			Call Restore()
		Case Else
			Addads()
	End Select

End Sub

Sub Page_main1()
%>
<style type="text/css">
table {width:100%;}
td {padding-left:5px;}
</style>
<table cellspacing="0" cellpadding="0" class="pw_tb1">
<tr><th>��Ϸ����˵��</th></tr>
<tr><td>
<ul>
<li>˵���٣����λ����Ψһ�������ظ���</li>
<li>˵���ڣ���ע��<font color="green">ѡ����ҳ������ʾ��λ��</font>��</li>
</ul>
</td></tr>
<tr><td>
<a href="?act=addads">�������λ</a>
| <a href="?act=adslist">�ҵĹ��λ</a>
| <a href="?act=restore">������</a>
</td></tr>
</table><br/>
<%
End Sub

'���/�༭���λ
Sub Addads()
%>
<style>
#adsize_label span{color:red;}
</style>
<script language="JavaScript">
function _setdemo(v){
	var obj = document.getElementById('ads_setdemo');
	if (obj){
		obj.src = "<%=Plus_Popwan.Folder%>images/ads_set"+v+".gif";
	}
}
</script>

<table cellspacing="0" cellpadding="0" class="pw_tb1">
<form method="post" action="?act=saveads" name="adsform">
<tr><th colspan="2" style="text-align:center;">�������λ��Ϣ(<font class="font2">����Ϊ������</font>)</th></tr>
<tr><td align="right">���λ��Ϣ��ʾ��</td>
<td width="85%"><div id="adsshow"></div></td>
</tr>
<tr>
<td width="15%" align="right">���λ���ƣ�</td>
<td width="85%">
<input type="text" name="zonename" size="20" value=""/>(���磺xxx��վ�������λ) 
<br/>
<font class="font1"></font>
</td>
</tr>
<tr>
<td  align="right">�����ʽ��</td>
<td>
<input type="radio" name="format" class="radio" value="1"/>���ֹ��  <input type="radio" name="format" checked="true" class="radio"  value="2"/>ͼƬ���
</td>
</tr>
<tr>
<td  align="right">�����룺</td>
<td>
<textarea name="zonedesc" style="width:96%;height:80px;"></textarea>
<br/><button name="pw" onclick="window.open('http://union.popwan.com/my/site/spread/<%=Plus_Popwan.ConfigNode.getAttribute("siteid")%>','popwanads');">��ȡ������</button>
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
<font class="green">ѡ����ҳ������ʾ��λ��</font>
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
<img id="ads_setdemo" src="<%=Plus_Popwan.Folder%>images/ads_set.gif" border="0" alt="���λ��Ԥ��"/>
</div>
</td>
</tr>
<tr><td colspan="2" align="center">
<input type="submit" name="submit" value="�����ҵĹ��λ"/>
</td></tr>
</form>
</table>
<%
End Sub

'������λ��Ϣ
Sub SaveAds()
	Dim zonename,format,adsize,getboard,adsset,zonedesc
	Dim homepage
	zonename = Trim(Request.Form("zonename")) '���λ����
	format = Trim(Request.Form("format")) '�����ʽ
	getboard = Trim(Request.Form("getboard")) 'ѡ����ʾ�İ���
	adsset = Trim(Request.Form("adsset"))
	zonedesc  = Trim(Request.Form("zonedesc")) '������
	'�ύ��Ϣ��֤
	If zonename=""or Len(zonename)<1 or Len(zonename)>32 Then
		Errmsg=ErrMsg + "<BR/>���λ���Ʋ���Ϊ�ջ򳬹�32���ַ���"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
	If format="" or Not IsNumeric(format) Then
		Errmsg=ErrMsg + "<BR/>����ȷѡ������ʽ��"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
	If zonedesc="" Then
		Errmsg=ErrMsg + "<BR/>�����벻��Ϊ�գ�"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
	If adsset = "" or Not Isnumeric(adsset) Then
		Errmsg=ErrMsg + "<BR/>��ѡȡ�����ҳ������ʾ��λ�ã�"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
	If format = "1" and (adsset="3" or adsset="4") Then
		Errmsg=ErrMsg + "<BR/>���ֹ�����ͣ����������ڸ��������¹̶����λ��"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If

	'�Ƿ���ʾ����ҳ����
	homepage = Request.Form("getskinid") '����ҳ:1 ��̳Ĭ�Ϲ��

		Set AdsList = popwan_ads.selectNodes("ads[@name='"&zonename&"']")
		If AdsList.Length>0 Then
			Errmsg=ErrMsg + "<BR/>���λ���Ʋ������ظ���"
			Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
			Exit Sub
		End If
		Set AdsList = popwan_ads.selectNodes("ads[@adsset='"&adsset&"' and @getboard='"&getboard&"' ]")
		If AdsList.Length>0 Then
			Errmsg=ErrMsg + "<BR/>�Ѿ��и���ͬ���õĹ����Ϣ���벻Ҫ�ظ��������ã�"
			Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
			Exit Sub
		End If

		'����һ���µĹ�����ݽڵ�
		Set AdsList = popwan_ads.appendChild(Forum_Api.createNode(1,"ads",""))
		AdsList.setAttribute "createtime",Now()
		AdsList.setAttribute "updatetime",Now()

	AdsList.setAttribute "name",zonename
	AdsList.setAttribute "format",format
	AdsList.setAttribute "homepage",homepage
	AdsList.setAttribute "adsset",adsset
	AdsList.setAttribute "getboard",getboard
	AdsList.setAttribute "zonedesc",zonedesc

		Update_Forum_Api()
		UpdateAdsSeting()
		Dvbbs.Dvbbs_suc("�������λ�ɹ���")
End Sub

'�ҵĹ��λ
Sub MyAdsList()
	Set AdsList = popwan_ads.selectNodes("ads")
	If AdsList.Length=0 Then
		Errmsg=ErrMsg + "<BR/>��δ�й��λ���ݣ�"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
	Dim Node
%>
<br/>
<table cellspacing="0" cellpadding="0" class="pw_tb1">
<tr><th colspan="7" style="text-align:center;" colspan="4">�ҵĹ��λ��Ϣ</th></tr>
<tr>
<td class="td2 title" >���λ����</td>
<td class="td2 title" width="10%">���λ��ʽ</td>
<td class="td2 title" width="15%">����ʱ��</td>
<td class="td2 title" width="15%">����ʱ��</td>
<td class="td2 title" >����</td>
</tr>
<%
	For Each Node In AdsList
%>
<tr>
<td><%=Node.getAttribute("name")%></td> 
<td><%=adssettype(Node.getAttribute("adsset"))%></td>
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
<td><a href="?act=editads&name=<%=Node.getAttribute("name")%>">�༭</a></td>
</tr>
<%
	Next
%>
</table>
<%
Set AdsList = Nothing
End Sub
Function adssettype(num)
	Select Case num
		Case "1" : adssettype = "ҳ�涥��"
		Case "2" : adssettype = "ҳ��ײ�"
		Case "7" : adssettype = "����¥������"
		Case "8" : adssettype = "����¥�����"
		Case "9" : adssettype = "����¥���ұ�"
		Case "10" : adssettype = "����¥���ײ�"
		Case Else
			adssettype = "δ֪"
	End Select
End Function

'�༭���λ
Sub Editads()
	Dim adzoneid
	Adzoneid = Request.QueryString("name")
	If Adzoneid<>"" Then
		Set AdsList = popwan_ads.selectSingleNode("ads[@name='"&Adzoneid&"']")
		If AdsList is Nothing  Then
			Errmsg=ErrMsg + "<BR/>����Ҫ�༭�Ĺ��λ�����ڣ�"
			Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
			Exit Sub
		End If
	End If
%>
<style>
#adsize_label span{color:red;}
</style>
<table cellspacing="0" cellpadding="0" class="pw_tb1">
<form method="post" action="?act=saveeditads" name="adsform">
<tr><th colspan="2" style="text-align:center;">�༭���λ��Ϣ(<font class="font2">����Ϊ������</font>)</th></tr>
<tr><td align="right">���λ��Ϣ��ʾ��</td>
<td width="85%"><div id="adsshow"></div></td>
</tr>
<tr>
<td width="15%" align="right">���λ���ƣ�</td>
<td width="85%">
<input type="text" name="zonename" size="20" value="<%=AdsList.getAttribute("name")%>" disabled/>(���磺xxx��վ�������λ) 
<br/>
<font class="font1"></font>
</td>
</tr>
<tr>
<td  align="right">�����ʽ��</td>
<td>
<input type="radio" name="format" class="radio" value="1"/>���ֹ��  <input type="radio" name="format" checked="true" class="radio"  value="2"/>ͼƬ���
</td>
</tr>
<tr>
<td  align="right">�����룺</td>
<td>
<textarea name="zonedesc" style="width:96%;height:80px;"><%=AdsList.getAttribute("zonedesc")%></textarea>
<br/><button name="pw" onclick="alert('����Ϸ������˴���ҳ�档');">��ȡ������</button>
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
<font class="green">ѡ����ҳ������ʾ��λ��</font>
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
<img id="ads_setdemo" src="<%=Plus_Popwan.Folder%>images/ads_set.gif" border="0" alt="���λ��Ԥ��"/>
</div>
</td>
</tr>
<tr><td colspan="2" align="center">
<input type="submit" name="submit" value="ȷ��"/>
</td></tr>
</form>
</table>
<script language="JavaScript">
function _setdemo(v){
	var obj = document.getElementById('ads_setdemo');
	if (obj){
		obj.src = "<%=Plus_Popwan.Folder%>images/ads_set"+v+".gif";
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
	Dim zonename,format,adsize,getboard,adsset,zonedesc
	Dim homepage
	zonename = Trim(Request.Form("zonename")) '���λ����
	format = Trim(Request.Form("format")) '�����ʽ
	getboard = Trim(Request.Form("getboard")) 'ѡ����ʾ�İ���
	adsset = Trim(Request.Form("adsset"))
	zonedesc  = Trim(Request.Form("zonedesc")) '������
	'�ύ��Ϣ��֤
	If format="" or Not IsNumeric(format) Then
		Errmsg=ErrMsg + "<BR/>����ȷѡ������ʽ��"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
	If zonedesc="" Then
		Errmsg=ErrMsg + "<BR/>�����벻��Ϊ�գ�"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
	If adsset = "" or Not Isnumeric(adsset) Then
		Errmsg=ErrMsg + "<BR/>��ѡȡ�����ҳ������ʾ��λ�ã�"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
	If format = "1" and (adsset="3" or adsset="4") Then
		Errmsg=ErrMsg + "<BR/>���ֹ�����ͣ����������ڸ��������¹̶����λ��"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If

	'�Ƿ���ʾ����ҳ����
	homepage = Request.Form("getskinid") '����ҳ:1 ��̳Ĭ�Ϲ��

	Set AdsList = popwan_ads.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
	If AdsList is Nothing  Then
		Errmsg=ErrMsg + "<BR/>����Ҫ�༭�Ĺ��λ�����ڣ�"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		Exit Sub
	End If
		AdsList.setAttribute "updatetime",Now()
	AdsList.setAttribute "format",format
	AdsList.setAttribute "homepage",homepage
	AdsList.setAttribute "adsset",adsset
	AdsList.setAttribute "getboard",getboard
	AdsList.setAttribute "zonedesc",zonedesc
	Update_Forum_Api()
	UpdateAdsSeting()
	Dvbbs.Dvbbs_suc("���λ���޸ĳɹ���")
End Sub


'������
Sub Restore()
	Set Forum_Api = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Forum_Api.LoadXml("<forum_api/>")
	Set popwan_ads = Forum_Api.documentElement.appendChild(Forum_Api.createNode(1,"popwan_ads",""))
	popwan_ads.setAttribute "memberid",""
	popwan_ads.setAttribute "email",""
	popwan_ads.setAttribute "password",""
	popwan_ads.setAttribute "nickname",""
	popwan_ads.setAttribute "webname",""
	Update_Forum_Api()
End Sub

Sub ChkForum_api()
	Set Forum_Api = Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not Forum_Api.load(NewsConfigFile) Then
		Creat_Forum_Api()
	Else
		Set Forum_Api = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Forum_Api.load(NewsConfigFile)
		Set popwan_ads = Forum_Api.documentElement.selectSingleNode("popwan_ads")
	End If
End Sub

Sub Creat_Forum_Api()
	Set Forum_Api = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Forum_Api.LoadXml("<forum_api/>")
	Set popwan_ads = Forum_Api.documentElement.appendChild(Forum_Api.createNode(1,"popwan_ads",""))
	popwan_ads.setAttribute "memberid",""
	popwan_ads.setAttribute "email",""
	popwan_ads.setAttribute "password",""
	popwan_ads.setAttribute "nickname",""
	popwan_ads.setAttribute "webname",""
	Update_Forum_Api()
End Sub

Sub Update_Forum_Api()
	Forum_Api.save NewsConfigFile
	Set Forum_Api=Nothing
End Sub


'���ù��λ������ҳ����ʾ��λ��
Sub UpdateAdsSeting()
	Dim iSetting,i,Forum_ads
	Dim adsset,adsstr
	Dim Sql
	
	'�������ִ�
	'mm_alimama��Աid_��վid_���λid
	adsset = AdsList.getAttribute("adsset")
	Adsstr = AdsList.getAttribute("zonedesc")

	Adsstr = Replace(Adsstr,"$","") '����δ����
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
%>