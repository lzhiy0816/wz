<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.LoadTemplates("")

If Request("Action")="readme" Then
	Main()
ElseIf Request("Action")="wappush" Then
	WapPush()
Else
	MyRedirect()
End If

Sub Main()
	Dvbbs.stats="WAP�ƶ���̳˵��"
	Dvbbs.Nav()
	Dvbbs.Head_var 0,0,"WAP�ƶ���̳˵��","wap.asp?Action=readme"
	Dvbbs.Showerr()
	Get_ChallengeWord
	Dim Rs
	Set Rs=Dvbbs.Execute("Select Top 1 * From Dv_ChallengeInfo")
	If Rs.Eof And Rs.Bof Then
		Response.redirect "showerr.asp?ErrCodes=<li>��������ݣ�����ϵ������̳�ٷ������&action=OtherErr"
		Exit Sub
	End If
%>
<table border="0" cellpadding=3 cellspacing=1 align=center class=Tableborder1>
	<tr>
	<th height=23>WAP�ƶ���̳������̳����������ʱ��صĹ�ͨ</th></tr>
	<FORM METHOD=POST ACTION="http://bbs.ray5198.com/bbsapp/midwappush.jsp">
	<input type=hidden name="mouseid" value="<%=Rs("D_UserName")%>">
	<input type=hidden name="forumname" value="<%=Dvbbs.Forum_Info(0)%>">
	<input type=hidden name="forumid" value="<%=Rs("D_ForumID")%>">
	<input type=hidden name="forumurl" value="<%=Dvbbs.Get_ScriptNameUrl()%>wap.asp">
	<input type=hidden name="backurl" value="<%=Dvbbs.Get_ScriptNameUrl()%>wap.asp?Action=wappush">
	<input type=hidden name="type" value="<%
	If Request("t")<>"" Then
		Response.Write "1"
	Else
		If Request("s")<>"" Then
			Response.Write Request("s")
		Else
			Response.Write "0"
		End If
	End If
	%>">
	<input type=hidden name="seqno" value="<%=Session("challengeWord")%>">
	<input type=hidden name="imgurl" value="<%
	If InStr(Lcase(Request("t")),"showimg") Then
		Response.Write Request("t") & "&filename="&Request("filename")
	Else
		Response.Write Request("t")
	End If
	%>">
	<input type=hidden name="provider" value="1">
	<tr>
	<td height=23 class=Tablebody1 style="line-height: 18px">
	<BLOCKQUOTE><BR>
	<center><%=Dvbbs.Forum_Info(0)%> Ŀǰ�ѿ�ͨ�ֻ��ƶ���̳���ƶ����ͣ����ܣ������������������龰��������ʱ��صĹ�ͨ</center>
	<hr size=1 color=#000000>
	<B>�ֻ����ʵ�ַ</B>��<font color=red><%=Dvbbs.Get_ScriptNameUrl()%>wap.asp</font>��<font color=red><%=Replace(Lcase(Dvbbs.Get_ScriptNameUrl()),"http://","wsp://")%>wap.asp</font><BR>
	˵����ֻҪ����ͨ���ƶ�GPRS����ͨCDMA�����ݷ�����֧��WAP���ֻ��ϼ���ͨ������ַ���� <%=Dvbbs.Forum_Info(0)%><BR><BR>
	<%If Request("t")="" Then%><B>WAP PUSH�·�<font color=blue>���</font>��̳��ַ</B><%Else%>WAP PUSH�·������̳ͼƬ��ַ����ͨ������ת�����ֻ���<%End If%>�������������ֻ����룺<input type=text size=20 name="mobile" value="13">&nbsp;<input type=submit name=submit value="<%
	If Request("t")<>"" Then
		Response.Write "����ͼƬ���ֻ�"
	Else
		If Request("s")<>"" Then
			Response.Write "��ѷ������ͼ�����Ϸ"
		Else
			Response.Write "��ѷ�����̳��ַ"
		End If
	End If%>">
	<BR><BR>
	<%
	If Request("t")<>"" Then
		If InStr(Lcase(Request("t")),"showimg") Then
			Response.Write "<div align=center><img src="&Request("t")&"&filename="&Request("filename")&" onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></div><BR><BR>"
		Else
			Response.Write "<div align=center><img src="&Request("t")&" onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></div><BR><BR>"
		End If
	End If
	%>
	<B>��Ҫ�ֻ���̳Ӧ�÷�Χ</B>��<BR>
	&nbsp;&nbsp;&nbsp;&nbsp;�ֻ�Ҳ�����͡���ͨ���ֻ���������ʱ��ص��������������̳�ϵĻ��Ϣ���˽�����������ӣ��ռǣ����Ķ����ظ��������������ʱ��صĽ��������⡢����Ȥ�¡�Ȥζ��Ƭ���ֻ����գ��ȵ���Ϣͨ���ֻ����������߹�ͨ��������֡�<BR>
	&nbsp;&nbsp;&nbsp;&nbsp;��Ϣ�ġ���ȡ������Ҫ���ϲ�������Я������������Я����ô�죿����ֻҪ���� <%=Dvbbs.Forum_Info(0)%> �ϱ�����������ϣ�ͨ���ֻ����Ϳ���ʱ��ص����·��Ļ������ص�������Ϣ��<BR>
	&nbsp;&nbsp;&nbsp;&nbsp;����ġ�����������������������Լ����ܽ����û�д�ͳ�����豸�ɹ�������ѯ��������ֻҪͨ���ֻ������ŵ������̳��������������⣬���źܿ���ܵõ��������ѵĻش�<BR><BR>
	<B>���ͼƬ������</B><BR>
	����Ϊ����ṩ���������ͼƬ������ֻҪ������ȷ��ʾ���շѵ���Ŀ��ȫ��������ѵġ�����Ҫ���Ѵ�ң����ֻ������й��ƶ���Ҫ��ȡGPRS�����ѵģ��ο�����ÿK3��Ǯ����Ȼ�����ƶ����в�ͬ���ײͻ����Ż����ߣ�ֻ���鷳�����ѯ���ص�1860�ƶ��ͷ�̨ȥ�˽�ϸ���ˣ�<a href="?Action=readme&s=2"><font color=blue>���ͼƬ����</font></a><BR><BR>
	<B>�����Ϸ���أ�</B><BR>
	����Ϊ����ṩ����������ֻ���Ϸ��ֻҪ������ȷ��ʾ���շѵ���Ŀ��ȫ��������ѵġ�����Ҫ���Ѵ�ң����ֻ������й��ƶ���Ҫ��ȡGPRS�����ѵģ��ο�����ÿK3��Ǯ����Ȼ�����ƶ����в�ͬ���ײͻ����Ż����ߣ�ֻ���鷳�����ѯ���ص�1860�ƶ��ͷ�̨ȥ�˽�ϸ���ˣ�<a href="?Action=readme&s=2"><font color=blue>�����Ϸ����</font></a><BR><BR>	<B>�ֻ������ʷ�˵��</B>���ƶ�GPRS����ͨCDMA�����ݷ�����»��������ã��ο�����ÿK3��Ǯ����Ȼ�����ƶ����в�ͬ���ײͻ����Ż����ߣ�ֻ���鷳�����ѯ���ص�1860�ƶ��ͷ�̨ȥ�˽�ϸ���ˣ��������κζ���ķ���֧��<BR><BR>
	<B>���������Դ��ע�����</B><BR>
	����������ͼƬ������������Ϸ����Ҫ������ȷ����Ļ����ǲ��Ǻ��ṩ���ļ�ƥ�䣬��Ϊ�ֻ��Ļ��Ͳ�ͬ������׼Ҳ��ͬ�����Ը��ֳ����ͨ���Ժܲһ��Ҫ�Ժ�ʹ�ã���Ȼ�������ص��ļ�����ʹ�ã���Ҫ�ķ������ص������ѣ���ĺ����۰���
	</BLOCKQUOTE>
	</td>
	</tr>
	</FORM>
</table>
<%
	Rs.Close
	Set Rs=Nothing
	Dvbbs.ActiveOnline()
	Dvbbs.Footer()
End Sub

Sub WapPush()
	Dvbbs.stats="WAP PUSH�·�"
	Dvbbs.Nav()
	Dvbbs.Head_var 0,0,"WAP PUSH�·�","wap.asp?Action=readme"
	Dvbbs.Showerr()
%>
<table border="0" cellpadding=3 cellspacing=1 align=center class=Tableborder1 style="width:600;">
	<tr>
	<th height=23>WAP�ƶ���̳������̳����������ʱ��صĹ�ͨ</th></tr>
	<tr>
	<td height=23 class=Tablebody1 style="line-height: 18px">
<%
	If Request("errorcode")="1" Then
		Response.Write "<li>���ɹ��Ľ���ص���̳��ַ��WAP PUSH�·���ָ�����ֻ���"
	Else
		Response.Write "<li>����WAP PUSH�·�ʧ�ܣ�"&request("errormsg")&"��"
	End If
%>
	</td>
	</tr>
</table>
<%
	Dvbbs.ActiveOnline()
	Dvbbs.Footer()
End Sub

Sub MyRedirect()
	Dim ServerUrl,ScriptName
	Dim Tmpstr
	Tmpstr = Request.ServerVariables("PATH_INFO")
	Tmpstr = Split(Tmpstr,"/")
	ScriptName = Lcase(Tmpstr(UBound(Tmpstr)))
	
	If request.servervariables("SERVER_PORT")="80" Then
		ServerUrl="http://" & request.servervariables("server_name")&replace(lcase(request.servervariables("script_name")),ScriptName,"")
	Else
		ServerUrl="http://" & request.servervariables("server_name")&":"&request.servervariables("SERVER_PORT")&replace(lcase(request.servervariables("script_name")),ScriptName,"")
	End If
	'Response.Write "http://bbs.ray5198.com/wap/first.jsp?url="&ServerUrl
	'Response.End
	Response.Redirect "http://bbs.ray5198.com/wap/first.jsp?url="&ServerUrl
End Sub
%>