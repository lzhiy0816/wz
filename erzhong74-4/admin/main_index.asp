<!--#include file="../Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Head()
CheckAdmin(",")
Page_Main()
Footer()

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Dvbbs.iCreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function CheckObj(objid)
	If Not IsObjInstalled(objid) Then
		CheckObj = "<font color="""&dvbbs.mainsetting(1)&""">��</font>"
	Else
		CheckObj = "��"
	End If
End Function

Sub Page_Main()
	Dim theInstalledObjects(20)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"	'Jamil 4.2
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
	theInstalledObjects(18) = "JMail.Message"	'Jamil 4.3
	theInstalledObjects(19) = "Persits.Upload"
	theInstalledObjects(20) = "SoftArtisans.FileUp"
	
	Dim Rs
	Dim Isaudituser
	Set Rs=Dvbbs.Execute("select count(*) from [dv_user] where usergroupid=5")
	Isaudituser=rs(0)
	if isnull(isaudituser) then isaudituser=0

	Dim BoardListNum
	set rs=dvbbs.execute("select count(*) from dv_board")
	BoardListNum=rs(0)
	If isnull(BoardListNum) then BoardListNum=0

%>
<table cellpadding="3" cellspacing="1" border="0" width="100%" align=center>
<tr><td colspan=2 height=25 class="td_title">��̳��Ϣͳ��</td></tr>
<tr><td height=23 colspan=2>

ϵͳ��Ϣ����̳������ <B><%=Dvbbs.CacheData(8,0)%></B> ������ <B><%=Dvbbs.CacheData(7,0)%></B> �û��� <B><%=Dvbbs.CacheData(10,0)%></B> ������û��� <B><%=Isaudituser%></B> �������� <B><%=BoardListNum%></B>

</td></tr>
<tr><td  class="forumRowHighlight" height=23 colspan=2>
����̳�ɶ����ȷ棨Dvbbs.Net����Ȩ�� <%=Dvbbs.Forum_info(0)%> ʹ�ã���ǰʹ�ð汾Ϊ ������̳
<%
If IsSqlDatabase=1 Then
	Response.Write "SQL���ݿ�"
Else
	Response.Write "Access���ݿ�"
End If
Response.Write " Dvbbs " & Dvbbs.Forum_Version
%>
</td></tr>
<tr>
<td width="50%"  height=23>���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRow">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="40%" height=23>
FSO�ı���д��<b><%=CheckObj(theInstalledObjects(9))%></b>

</td>
<td width="60%" class="forumRow">
Jmail4.3�������֧�֣�
<b><%=CheckObj(theInstalledObjects(18))%></b>
&nbsp;&nbsp;<a href="data.asp?action=SpaceSize">>>�鿴����ϸ��������Ϣ���</a>
</td>
</tr>
<tr><td height=23 colspan=2>
<a href="challenge.asp"><font color=red>����֧���������</font></a>����Ҫ������̳��ȯ��ֵ����ص�����̳����ķ��񣬵�ǰ֧���ʺ�Ϊ<%=Dvbbs.Forum_ChanSetting(4)%>
</td></tr>
<tr><td height=23 colspan=2>
<a href="data.asp"><font color=red>���ݶ��ڱ���</font></a>����ע�����ö������ݱ��ݣ����ݵĶ��ڱ��ݿ�����޶ȵı�������̳���ݵİ�ȫ
</td></tr>
</table><br />


<%
Dim Forum_Pack
Set Rs=Dvbbs.Execute("Select Forum_Pack From Dv_Setup")
Forum_Pack = Rs(0)
Forum_Pack=Split(Forum_Pack,"|||")
If UBound(Forum_Pack)<2 Then ReDim Forum_Pack(3)
If Forum_Pack(0) = "1" Then
%>
<table cellpadding="3" cellspacing="1" border="0" align=center width="100%">
<tr><th colspan="2"><a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=143" target="SiteAdmin">��̳������</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=134" target="SiteAdmin">��̳������</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=8" target="SiteAdmin">��̳������</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=13" target="SiteAdmin">��̳�����</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=102" target="SiteAdmin">��̳�����</a></td><tr>
<tr><td height=23 valign=top>
<iframe src="http://bbs.dvbbs.net/union_post.asp" name="SiteAdmin" height=180 width="100%" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=no></iframe>
</tr>
</table>
<%
End If
Rs.Close
Set Rs=Nothing
%>

<br />
<table cellpadding="3" cellspacing="1" border="0" align=center>
<tr><td class="td_title" colspan=2 height=25>��̳����С��ʿ</td></tr>
<tr><td height=23 align="center" width="15%">
<img src="skins/images/user_tag.gif">
</td><td height=23 width="*"><B>�û���Ȩ��</B><br />������̳��ע���û��ֳɲ�ͬ���û��飬ÿ���û������ӵ�в�ͬ����̳����Ȩ�ޣ������ڶ�����̳7.0�汾֮���û��ȼ���ϵ����û����У������û��ȼ�û���Զ���Ȩ�ޣ���ô����ȼ���Ȩ�޾�ʹ�����������û���Ȩ�ޣ���֮��ӵ������ȼ��Լ���Ȩ�ޡ�<font color=red>ÿ���ȼ������û������趨��Ȩ�޶��������������̳��</font></td></tr>
<tr><td height=23 align="center">
<img src="skins/images/tag.gif">
</td><td height=23 width="*"><B>�ְ���Ȩ��</B><br />
ÿ���û�������Զ���Ȩ�����õĵȼ�������������������̳�и�������ӵ�в�ͬ��Ȩ�ޣ�����˵����������ע���û�����������·�ڰ���A���ܷ�����������ȵ�Ȩ�����ã��������������̳Ȩ�޵����ã�<font color=blue>����������˵���Էֳ��ܶ����ͬ�������͵���̳</font>
</td></tr>
<tr><td height=23 align="center">
<img src="skins/images/user_sys.gif">
</td><td height=23 width="*"><B>�û�Ȩ���趨</B><br />
ÿ���û�����������������̳�и�������ӵ�в�ͬ��Ȩ�޻��������Ȩ�ޣ�����˵�����������û�A�ڰ���A��ӵ�����й���Ȩ�ޡ�<U>������������Ȩ����Ҫע�����������˳��Ϊ���û�Ȩ������(<font color=gray>�Զ���</font>)<font color=blue> <B>></B> </font>�ְ���Ȩ���趨(<font color=gray>�Զ���</font>)<font color=blue> <B>></B> </font>�û���Ȩ���趨(<font color=gray>Ĭ��</font>)</U>
</td></tr>
<tr><td height=23 align="center">
<img src="skins/images/skins.gif">
</td><td height=23 width="*"><B>�Է��ģ��Ĺ���</B><br />
���а�������̳����ģ��Ĺ���ģ������̳�Ļ���CSS���ã���̳�����ĸ��ģ���̳��ҳ����ĸ��ģ�ͼƬ�����ã����԰������ã��½�ģ��ҳ���ģ�壬ģ�����½���ͬ�����ԡ�ͼƬ������ģ��Ԫ�صȵȹ��ܣ�����ӵ��ģ��ĵ��뵼�����ܣ�������������ʵ������̳�������߱༭���л�
</td></tr>
<tr><td height=23 valign=top  align="right">
<B>һ�仰��ʿ</B>
</td><td height=23 width="*"><B>һ�仰��ʿ</B><br />
�� ���ڲ�ͬ����ģ���ҳ�棬Ҫ��ϸ��ҳ���е�˵�������������
<BR>
�� �û��鼰����չ��Ȩ�����ã�����̳�ĸ��������м��������ԣ�Ҫ������������Ⱥ���Ч˳��
<BR>
�� �����̳������ʱ�򣬱����˻�ͷ�����ð���߼������Ƿ���ȷ
<BR>
�� �������뵽������̳�ٷ�վ�����ʣ��кܶ����ĵ����ѻ��æ��<a href="help.asp">�鿴������ʿ����</a>
</td></tr>
</table><br />

<table width="100%" border="0" cellspacing="1" cellpadding="0">
    <form name="form1" method="post" action="">
  <tr>
    <td colspan="2" valign="middle" class="td_title">��̳�������</td>
  </tr>
  <tr>
    <td align="right" width="15%">��Ʒ������</td>
    <td><a href="#01">���ڶ����ȷ�����Ƽ����޹�˾</a>  �й����Ұ�Ȩ������Ȩ�ǼǺ�2004SR00001 </td>
  </tr>
  <tr>
    <td align="right">��Ʒ����</td>
    <td>��վ��ҵ�� ������̳��Ŀ��  ��ҵ���Ͱ��� </td>
  </tr>
  <tr>
    <td align="right">��ϵ��ʽ��</td>
    <td>��վ��ҵ����0898-68557467<br />
������ҵ����0898-68592224 68592294<br />
���������棺0898-68556467<br />
����鿴��ϸ��ϵ���� </td>
  </tr>
  <tr>
    <td align="right">���������</td>
    <td>
      ������̳�����֯��Dvbbs Plus Organization�� </td>
  </tr>
    </form>
</table>
<%
End Sub
%>