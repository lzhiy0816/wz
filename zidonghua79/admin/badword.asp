<!--#include file=../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
	Head()
	dim admin_flag
	Dim XMLDom,node,nodename,checkedstr
	admin_flag="26,27"
	if not Dvbbs.master or instr(","&session("flag")&",",",26,")=0 or instr(","&session("flag")&",",",27,")=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		call main()
	Footer()
	end if

	sub main()
	dim sel

if request("action") = "savebadword" then
call savebadword()
else

%>

<form action="badword.asp?action=savebadword" method=post>

<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">

<%if request("reaction")="badword" then%>
<tr>
<th colspan=2 align=left height=23>���ӹ����ַ�</th>
</tr>
<tr>
<td class=forumrow width="100%" colspan=2><B>˵��</B>�������ַ��趨����Ϊ  <B>Ҫ���˵��ַ�=���˺���ַ�</B> ��ÿ�������ַ��ûس��ָ��</td>
</tr>
<tr>
<td class=forumrow width="100%" colspan=2>
<textarea name="badwords" cols="80" rows="8"><%
For i=0 To Ubound(Dvbbs.BadWords)
	If i > UBound(Dvbbs.rBadWord) Then
		Response.Write Dvbbs.BadWords(i) & "=*"
	Else
		Response.Write Dvbbs.BadWords(i) & "=" & Dvbbs.rBadWord(i)
	End If
	If i<Ubound(Dvbbs.BadWords) Then Response.Write chr(10)
Next
%></textarea>
</td>
</tr>
<%elseif request("reaction")="splitreg" Then
LoadSetting
	%>
<tr>
<th colspan=2 align=left height=23>ע������ַ�</th>
</tr>
<tr>
<td class=forumrow width="20%">˵����</td>
<td class=forumrow width="80%">ע������ַ����������û�ע����������ַ������ݣ�������Ҫ���˵��ַ������룬����ж���ַ��������á�,���ָ��������磺ɳ̲,quest,ľ��</td>
</tr>
<tr>
<td class=forumrow width="20%">����������ַ�</td>
<td class=forumrow width="80%"><input type="text" name="splitwords" value="<%=split(Dvbbs.cachedata(1,0),"|||")(4)%>" size="80"></td>
</tr>
</table>
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr>
<th colspan=2 align=left height=23>ע����������</th>
</tr>
<tr>
<td class=forumrow width="20%">˵����</td>
<td class=forumrow width="80%">��չע������,������Լ���Ҫ����.</td>
</tr>
<tr>
<%
nodename="checkip"
Set Node=XMLDom.documentElement.selectSingleNode(nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.selectSingleNode("@use").text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<td class=forumrow width="20%">����IP����</td>
<td class=forumrow width="80%"><input type="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %>  /> �����ѡ��,���к�IP��ַ�йص����ö���������.</td>
</tr>
<tr>
<td class=forumrow width="20%">����ע��IP(IP������)<br>
��д����ע���IP��ַ,��ʽ��IP��ַ=˵��,ÿ���û��зֿ�
֧��ͨ���,��192.168.*.* =����IP �粻����IP������,������
</td>
<td class=forumrow width="80%"><textarea name="iplist1" cols="80" rows="8"><%
For Each  Node In XMLDom.documentElement.selectNodes("checkip/iplist1/ip")
	Response.Write Node.text &" = "&Node.selectSingleNode("@description").text &Chr(10)
Next
%></textarea></td>
</tr>
<tr>
<td class=forumrow width="20%">��ֹע��IP(ip������)<br>
��д����ע���IP��ַ,��ʽ��IP��ַ=˵��,ÿ���û��зֿ�
֧��ͨ���,��192.168.*.* =����IP �粻����IP������,������
</td>
<td class=forumrow width="80%"><textarea name="iplist2" cols="80" rows="8"><%
For Each  Node In XMLDom.documentElement.selectNodes("checkip/iplist2/ip")
	Response.Write Node.text &" = "&Node.selectSingleNode("@description").text &Chr(10)
Next
%></textarea></td>
</tr>
<%
nodename="postipinfo"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<tr>
<td class=forumrow width="20%">��IP�������ύIP��Դ��Ϣ</td>
<td class=forumrow width="80%">
<input type="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %> /><br> ���ע���û�����IP��������ע��֮��,��������ע���߽����ύ��ǰIP��Ϣ��ҳ��,�Ա����Ա�������Ӹö�IP��ַ������.
</td>
</tr>
<%
nodename="checkregcount"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr="1"
Else
		checkedstr=Node.text
End If
%>
<tr>
<td class=forumrow width="20%">һ��IP��ַһ�����ע�����</td>
<td class=forumrow width="80%"><input type="text" Size="4" value="<%=checkedstr %>" name="<%=nodename%>"  />����д���ַ�������,�粻������,����д0</td>
</tr>
<%
nodename="checknumeric"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<tr>
<td class=forumrow width="20%">��ֹ������IDע��</td>
<td class=forumrow width="80%"><input type="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %> />�Ƿ��������ô������û���ע��</td>
</tr>
<%
nodename="checktime"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<tr>
<td class=forumrow width="20%">Ҫ�����뵱ǰʱ��</td>
<td class=forumrow width="80%"><Input type="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %> /><br>�������,Ҫ���û�ѡ���Լ�����ʱ�������������ڵص�ʱ��(��СʱΪ��λ)</td>
</tr>
<%
nodename="usevarform"
Set Node=XMLDom.documentElement.selectSingleNode("@"&nodename)
If Node Is Nothing Then
	checkedstr=" checked=""checked"""
Else
If Node.text="1" Then
	checkedstr=" checked=""checked"""
Else
	checkedstr=""
End If
End If
%>
<tr>
<td class=forumrow width="20%">���ö�̬�ı�����Ŀ����</td>
<td class=forumrow width="80%"><input type="checkbox" value="1" name="<%=nodename%>" <%=checkedstr %> />���ò������ı�����Ŀ����,���ӻ�����ע����Ѷ�.</td>
</tr>
<%end if%>
<input type=hidden value="<%=request("reaction")%>" name="reaction">
<tr> 
<td class=forumrow width="20%"></td>
<td width="80%" class=forumrow><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table>

</form>
<%end if%>
<%
end sub

sub savebadword()
dim iforum_setting,forum_setting
If request("reaction")="badword" then
dim badwords,badwords_1,badwords_2,badwords_3
badwords=request("badwords")
badwords=split(badwords,vbCrlf)
for i = 0 to ubound(badwords)
	if not (badwords(i)="" or badwords(i)=" ") then
		badwords_1 = split(badwords(i),"=")
		If ubound(badwords_1)=1 Then
			If i=0 Then
				badwords_2 = badwords_1(0)
				badwords_3 = badwords_1(1)
			Else
				badwords_2 = badwords_2 & "|" & badwords_1(0)
				badwords_3 = badwords_3 & "|" & badwords_1(1)
			End If
		End If
	End If
next

sql = "update dv_setup set Forum_Badwords='"&replace(badwords_2,"'","''")&"',Forum_rBadword='"&replace(badwords_3,"'","''")&"'"
dvbbs.execute(sql)
elseif request("reaction")="splitreg" then
Call LoadXML
'forum_info|||forum_setting|||forum_user|||copyright|||splitword|||stopreadme
Set rs=Dvbbs.execute("select forum_setting from dv_setup")
iforum_setting=split(rs(0),"|||")
forum_setting=iforum_setting(0) & "|||" & iforum_setting(1) & "|||" & iforum_setting(2) & "|||" & iforum_setting(3) & "|||" & request("splitwords") & "|||" & iforum_setting(5)
sql = "update dv_setup set forum_setting='"&replace(forum_setting,"'","''")&"',Forum_Boards='"&Dvbbs.checkstr(XMLDom.XML)&"'"
dvbbs.execute(sql)
End If
Dvbbs.loadSetup()
Dv_suc("���³ɹ�")

End Sub
Sub LoadSetting()
	Dim Rs,Nedupdate
	Set Rs=Dvbbs.Execute("Select Forum_Boards From Dv_setup")
	Set XMLDom=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not (xmldom.LoadXML(Rs(0))) Then
		Nedupdate=True
	ElseIf xmldom.documentElement.nodeName<>"regsetting" Then
		Nedupdate=True
	End If
	If Nedupdate Then
		XMLDom.LoadXML "<?xml version=""1.0""?><regsetting/>"
		Dvbbs.Execute"update Dv_setup Set Forum_Boards='"&Dvbbs.checkstr(XMLDom.XML)&"'"
	End If
End Sub
Sub LoadXML()
	Dim Node,node1,i,iplist,node2,iplist1
	Set XMLDom=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("regsetting"))
	Set Node=xmldom.documentElement.appendChild(XMLDom.createNode(1,"checkip",""))
	If Request.form("checkip")="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"use","")).text="1"
	Else
		Node.attributes.setNamedItem(XMLDom.createNode(2,"use","")).text="0"
	End If
	Set node1=Node.appendChild(XMLDom.createElement("iplist1"))
	For each iplist in split(Request.form("iplist1"),vbnewline)
		If iplist<>"" Then
			iplist1=Split(iplist,"=")
			If UBound(iplist1)>0 Then
			Set node2=node1.appendChild(XMLDom.createNode(1,"ip",""))
			node2.text=Trim(iplist1(0))
			Node2.attributes.setNamedItem(XMLDom.createNode(2,"description","")).text=Trim(iplist1(1))
			End If
		End If
	Next
	Set node1=Node.appendChild(XMLDom.createElement("iplist2"))
	For each iplist in split(Request.form("iplist2"),vbnewline)
		If iplist<>"" Then
			iplist1=Split(iplist,"=")
			If UBound(iplist1)>0 Then
			Set node2=node1.appendChild(XMLDom.createNode(1,"ip",""))
			node2.text=Trim(iplist1(0))
			Node2.attributes.setNamedItem(XMLDom.createNode(2,"description","")).text=Trim(iplist1(1))
			End If
		End If
	Next
	If Request.form("postipinfo")="1" Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"postipinfo","")).text="1"
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"postipinfo","")).text="0"
	End If
	'If Request.form("checkproxy")="1" Then
	'		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkproxy","")).text="1"
	'Else
	'	xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkproxy","")).text="0"
	'End If
	If Request.form("checknumeric")="1" Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checknumeric","")).text="1"
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checknumeric","")).text="0"
	End If
	If Request.form("checktime")="1" Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checktime","")).text="1"
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checktime","")).text="0"
	End If
	If Request.form("usevarform")="1" Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"usevarform","")).text="1"
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"usevarform","")).text="0"
	End If
	If Request.form("checkregcount")<>"" and IsNumeric(Request.form("checkregcount")) Then
			xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkregcount","")).text=Request.form("checkregcount")
	Else
		xmldom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkregcount","")).text="0"
	End If
	
End Sub
%>