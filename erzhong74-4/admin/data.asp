<!--#include file =../conn.asp-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim TestConn,action
Dim admin_flag
Dim dbpath,bkfolder,bkdbname,fso,fso1
Dim uploadpath
Dim okOS,okCpus,okCPU
action=Trim(request("action"))

If Dvbbs.Forum_Setting(76)="0" Or  Dvbbs.Forum_Setting(76)="" Then Dvbbs.Forum_Setting(76)="UploadFile/"
uploadpath="../"&Dvbbs.Forum_Setting(76)
Select Case action
	Case "SpaceSize"		'ϵͳ�ռ�ռ��
		admin_flag=",35,"
		CheckAdmin(admin_flag)
		Call SpaceSize()
	Case "CompressData","BackupData","RestoreData"
		Call ReadMe()
	Case Else
		Errmsg=ErrMsg + "<BR><li>ѡȡ��Ӧ�Ĳ�����"
		dvbbs_error()
End Select

Footer()
response.write"</body></html>"


'====================ϵͳ�ռ�ռ��=======================
sub SpaceSize()
On error resume next
GetSysInfo()
Dim t
't = GetAllSpace
Dim FoundFso
FoundFso = False
FoundFso = IsObjInstalled("Scripting.FileSystemObject")
%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%"><tr>
<th style="text-align:center;" colspan=5>
&nbsp;&nbsp;ϵͳ��Ϣ������
</th>
</tr>
<tr>
<td class="td1" width="35%" height=23>
��ǰ��̳�汾
</td>
<td class="td1" width="15%">
<a href="http://www.dvbbs.net/download.asp" target=_blank>Dvbbs <%=Dvbbs.Forum_Version%></a>
</td>
<td width="8" class="td1">&nbsp;</td>
<td class="td1" width="35%">
���ݿ����ͣ�
</td>
<td class="td1" width="15%">
<%
If IsSqlDataBase = 1 Then
	Response.Write "Sql Server"
Else
	Response.Write "Access"
End If
%>
</td>
</tr>
<tr>
<td class="td2" width="35%" height=23>
����������IP
</td>
<td class="td2" width="15%">
<%=Request.ServerVariables("SERVER_NAME")%><BR><%=Request.ServerVariables("LOCAL_ADDR")%>
</td>
<td width="8" class="td2">&nbsp;</td>
<td class="td2" width="35%">
���ݿ�ռ�ÿռ�
</td>
<td class="td2" width="15%">
<%
If IsSqlDataBase = 1 Then
	Set Rs=Dvbbs.Execute("Exec sp_spaceused")
	If Err <> 0 Then
		Err.Clear
		Response.Write "<font color=gray>δ֪</font>"
	Else
		Response.Write Rs(1)
	End If
Else
	If FoundFso Then
		Response.Write GetFileSize(MyDbPath & DB)
	Else
		Response.Write "<font color=gray>δ֪</font>"
	End If
End If
%>
</td>
</tr>
<tr>
<td class="td1" width="35%" height=23>
�ϴ�ͷ��ռ�ÿռ�
</td>
<td class="td1" width="15%">
<%showSpaceinfo("../uploadface")%>
</td>
<td width="8" class="td1">&nbsp;</td>
<td class="td1" width="35%">
�ϴ�ͼƬռ�ÿռ�
</td>
<td class="td1" width="15%">
<%showSpaceinfo(uploadpath)%>
</td>
</tr>
<tr>
<td class="td2" width="100%" height=23 colspan=5>
<B>�����������Ϣ</B>
</td>
</tr>
<tr>
<td class="td1" width="35%" height=23>
ASP�ű���������
</td>
<td class="td1" width="15%">
<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>
</td>
<td width="8" class="td1">&nbsp;</td>
<td class="td1" width="35%">
IIS �汾
</td>
<td class="td1" width="15%">
<%=Request.ServerVariables("SERVER_SOFTWARE")%>
</td>
</tr>
<tr>
<td class="td2" width="35%" height=23>
����������ϵͳ
</td>
<td class="td2" width="15%">
<%=okos%>
</td>
<td width="8" class="td2">&nbsp;</td>
<td class="td2" width="35%">
������CPU����
</td>
<td class="td2" width="15%">
<%=okcpus%> ��
</td>
</tr>
<tr>
<td class="td1" width="100%" height=23 colspan=5>
���ļ�·����<%=Server.Mappath("data.asp")%>
</td>
</tr>
<tr>
<td class="td2" width="100%" colspan=5 height=23>
<B>��Ҫ�����Ϣ</B>
</td>
</tr>
<tr>
<td class="td1" width="35%" height=23>
FSO�ļ���д
</td>
<td class="td1" width="15%">
<%
If FoundFso Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
<td width="8" class="td1">&nbsp;</td>
<td class="td1" width="35%">
Jmail�����ʼ�֧��
</td>
<td class="td1" width="15%">
<%
If IsObjInstalled("JMail.SmtpMail") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="td2" width="35%" height=23>
CDONTS�����ʼ�֧��
</td>
<td class="td2" width="15%">
<%
If IsObjInstalled("CDONTS.NewMail") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
<td width="8" class="td2">&nbsp;</td>
<td class="td2" width="35%">
AspEmail�����ʼ�֧��
</td>
<td class="td2" width="15%">
<%
If IsObjInstalled("Persits.MailSender") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="td1" width="35%" height=23>
������ϴ�֧��
</td>
<td class="td1" width="15%">
<%
If IsObjInstalled("Adodb.Stream") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
<td width="8" class="td1">&nbsp;</td>
<td class="td1" width="35%">
AspUpload�ϴ�֧��
</td>
<td class="td1" width="15%">
<%
If IsObjInstalled("Persits.Upload") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="td2" width="35%" height=23>
SA-FileUp�ϴ�֧��
</td>
<td class="td2" width="15%">
<%
If IsObjInstalled("SoftArtisans.FileUp") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
<td width="8" class="td2">&nbsp;</td>
<td class="td2" width="35%">
DvFile-Up�ϴ�֧��
</td>
<td class="td2" width="15%">
<%
If IsObjInstalled("DvFile.Upload") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="td1" width="35%" height=23>
CreatePreviewImage����Ԥ��ͼƬ
</td>
<td class="td1" width="15%">
<%
If IsObjInstalled("CreatePreviewImage.cGvbox") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
<td width="8" class="td1">&nbsp;</td>
<td class="td1" width="35%">
AspJpeg����Ԥ��ͼƬ
</td>
<td class="td1" width="15%">
<%
If IsObjInstalled("Persits.Jpeg") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="td2" width="35%" height=23>
SA-ImgWriter����Ԥ��ͼƬ
</td>
<td class="td2" width="15%">
<%
If IsObjInstalled("SoftArtisans.ImageGen") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
</td>
<td width="8" class="td2">&nbsp;</td>
<td class="td2" width="35%">ADO(���ݿ����)�汾:<%=conn.Version%>
</td>
<td class="td2" width="15%"><font color=green><b>��</b></font>
</td>
</tr>
<FORM action="data.asp?action=SpaceSize" method=post id=form1 name=form1>
<tr>
<td class="td1" width="100%" height=23 colspan=5>
<%
If Request("classname")<>"" Then
	If IsObjInstalled(Request("classname")) Then
		Response.Write "<font color=green><b>��ϲ����������֧�� "&Request("classname")&" ���</b></font><BR>"
	Else
		Response.Write "<font color=red><b>��Ǹ������������֧�� "&Request("classname")&" ���</b></font><BR>"
	End If
End If
%>
�������֧�������ѯ��<input class=input type=text value="" name="classname" size=30>
<INPUT type=submit class="button" value="�� ѯ" id=submit1 name=submit1>
��������� ProgId �� ClassId
</td>
</tr>
</form>
</table>
<%Response.Flush%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
<tr>
<td class="td2" width="100%" colspan=5 height=23>
<B>�����ļ������ٶȲ���</B>
</td>
</tr>
<tr>
<td class="td1" width="100%" colspan=5 height=23>
<%
	Response.Write "�����ظ�������д���ɾ���ı��ļ�50��..."

	Dim thetime3,tempfile,iserr,t1,FsoObj,tempfileOBJ,t2,i
	Set FsoObj=Dvbbs.iCreateObject("Scripting.FileSystemObject")

	iserr=False
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	For i=1 To 50
		Err.Clear

		Set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		If Err <> 0 Then
			Response.Write "�����ļ�����"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		If Err <> 0 Then
			Response.Write "д���ļ�����"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		If Err <> 0 Then
			Response.Write "ɾ���ļ�����"
			iserr=True
			Err.Clear
			Exit For
		end if
		Set tempfileOBJ=Nothing
	Next
	t2=timer
	If Not iserr Then
		thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
		Response.Write "...����ɣ���������ִ�д˲�������ʱ <font color=red>" & thetime3 & " ����</font>"
	End If
%>
</td>
</tr>
<tr>
<td class="td2" width="100%" height=23 colspan=5>
<a href="http://www.aspsky.cn" target=_blank>�����Ƽ��������� <font color=gray>˫��ǿ2.4,2GddrEcc,SCSI36.4G*2</font> ִ�д˲�����Ҫ <font color=red>32��65</font> ����</a>
</td>
</tr>
</table>
<%Response.Flush%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
<tr>
<td class="td1" width="100%" colspan=5 height=23>
<B>ASP�ű����ͺ������ٶȲ���</B>
</td>
</tr>
<tr>
<td class="td2" width="100%" colspan=5 height=23>
<%

	Response.Write "����������ԣ����ڽ���50��μӷ�����..."
	dim lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�����ʱ <font color=red>" & thetime & " ����</font><br>"


	Response.Write "����������ԣ����ڽ���20��ο�������..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�����ʱ <font color=red>" & thetime2 & " ����</font><br>"
%>
</td>
</tr>
<tr>
<td class="td1" width="100%" colspan=5 height=23>
<a href="http://www.aspsky.cn" target=_blank>�����Ƽ��������� <font color=gray>˫��ǿ2.4,2GddrEcc,SCSI36.4G*2</font> ����������Ҫ <font color=red>171��203</font> ����, ����������Ҫ <font color=red>156��171</font> ����</a>

</td>
</tr>
</table><BR>
<%
end sub



Sub ReadMe()
	If IsSqlDataBase=0 Then
		Call AccessUserReadme()
	Else
		Call SQLUserReadme()
	End If
End Sub

Sub AccessUserReadme()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th height=25 style="text-align:center;">
			&nbsp;&nbsp;Access���ݿ����ݴ���˵�� (����ǰ��ر���̳)
		</th>
	</tr> 	
	<tr>
		<td class="td1"> 			
		<blockquote>
			˵��������ȫԭ�򣬱��ݺͻ�ԭ���ݿ⹦����ȡ�������û����е�FTP�ϲ�����<br />
			�����������£�<br />
			<br /><b>���ݿⱸ�ݣ�</b><br />
			ʹ��FTP������¼FTPվ�㣬�������豸�ݵ����ݿ⣬���浽���ص��Դ�ţ�����ÿ�ܱ���һ�����ϡ�<br />
			<br /><b>���ݿ⻹ԭ��</b><br />
			���Ҫ��ԭ���ݿ⣬����֮ǰ�����ص����ݿ�,�滻ԭ�������ݿ⼴�ɡ�<br /> 
			<br /><b>ѹ���޸���</b><br />
			 ���ݿ��ļ���ʹ�ù����У��������ݵ���ɾ���ļ��ߴ������������Ҫ�������ѹ���޸���<br />
			 �������ݿ��ļ�[�����.asp����չ�������Ϊ.mdb����չ��]����Microsoft Access�����ݿ⣬ѡ�񹤾�--���ݿ�ʵ�ù���--ѹ�����޸����ݿ�--[�Ļ�.asp����չ��]--�ϴ�����ԭ�����ݿ��ļ�������ͼ��<br />

			 <b>Access2007</b> �޸�������ѡ��˵���ϵͳ���������ݿ���������ݿ��޸�ѹ����ϵͳ���Զ��޸�ѹ�����ݿ��ļ���<br /><br />
			 <b>Access2003</b> �޸������������ݿ����ص����أ���ACCESS 2003/xp �򿪺���ͼ��ʾѹ���޸�:<br />
			<img src="skins/images/access.gif" width="590" height="389" alt="" />
		</blockquote>
		</td>
	</tr>
</table>
<%
End Sub

sub SQLUserReadme()
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="100%">
				<tr>
  					<th height=25 style="text-align:center;">
  					&nbsp;&nbsp;SQL���ݿ����ݴ���˵��
  					</th>
  				</tr> 	
 				<tr>
 					<td class="td1"> 			
 			<blockquote>
<B>һ���������ݿ�</B>
<BR><BR>
1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
2��SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼<BR>
3��ѡ��������ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�������˵��еĹ���-->ѡ�񱸷����ݿ�<BR>
4������ѡ��ѡ����ȫ���ݣ�Ŀ���еı��ݵ����ԭ����·����������ѡ�����Ƶ�ɾ����Ȼ������ӣ����ԭ��û��·����������ֱ��ѡ�����ӣ�����ָ��·�����ļ�����ָ�����ȷ�����ر��ݴ��ڣ����ŵ�ȷ�����б���
<BR><BR>
<B>������ԭ���ݿ�</B><BR><BR>
1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
2��SQL Server��-->˫������ķ�����-->��ͼ�������½����ݿ�ͼ�꣬�½����ݿ����������ȡ<BR>
3������½��õ����ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�������˵��еĹ���-->ѡ��ָ����ݿ�<BR>
4���ڵ������Ĵ����еĻ�ԭѡ����ѡ����豸-->��ѡ���豸-->������-->Ȼ��ѡ����ı����ļ���-->���Ӻ��ȷ�����أ���ʱ���豸��Ӧ�ó������ղ�ѡ������ݿⱸ���ļ��������ݺ�Ĭ��Ϊ1���������ͬһ���ļ�������α��ݣ����Ե�����ݺ��ԱߵĲ鿴���ݣ��ڸ�ѡ����ѡ�����µ�һ�α��ݺ��ȷ����-->Ȼ�����Ϸ������Աߵ�ѡ�ť<BR>
5���ڳ��ֵĴ�����ѡ�����������ݿ���ǿ�ƻ�ԭ���Լ��ڻָ����״̬��ѡ��ʹ���ݿ���Լ������е��޷���ԭ����������־��ѡ��ڴ��ڵ��м䲿λ�Ľ����ݿ��ļ���ԭΪ����Ҫ������SQL�İ�װ�������ã�Ҳ����ָ���Լ���Ŀ¼�����߼��ļ�������Ҫ�Ķ������������ļ���Ҫ���������ָ��Ļ���������Ķ���������SQL���ݿ�װ��D:\Program Files\Microsoft SQL Server\MSSQL\Data����ô�Ͱ������ָ�������Ŀ¼������ظĶ��Ķ������������ļ�����øĳ�����ǰ�����ݿ�������ԭ����bbs_data.mdf�����ڵ����ݿ���forum���͸ĳ�forum_data.mdf������־�������ļ���Ҫ���������ķ�ʽ����صĸĶ�����־���ļ�����*_log.ldf��β�ģ�������Ļָ�Ŀ¼�������������ã�ǰ���Ǹ�Ŀ¼������ڣ���������ָ��d:\sqldata\bbs_data.mdf����d:\sqldata\bbs_log.ldf��������ָ�������<BR>
6���޸���ɺ󣬵�������ȷ�����лָ�����ʱ�����һ������������ʾ�ָ��Ľ��ȣ��ָ���ɺ�ϵͳ���Զ���ʾ�ɹ������м���ʾ���������¼����صĴ������ݲ�ѯ�ʶ�SQL�����Ƚ���Ϥ����Ա��һ��Ĵ����޷���Ŀ¼��������ļ����ظ������ļ���������߿ռ䲻���������ݿ�����ʹ���еĴ������ݿ�����ʹ�õĴ��������Գ��Թر����й���SQL����Ȼ�����´򿪽��лָ��������������ʾ����ʹ�õĴ�����Խ�SQL����ֹͣȻ�����𿴿����������������Ĵ���һ�㶼�ܰ��մ�����������Ӧ�Ķ��󼴿ɻָ�<BR><BR>

<B>�����������ݿ�</B><BR><BR>
һ������£�SQL���ݿ�����������ܴܺ�̶��ϼ�С���ݿ��С������Ҫ������������־��С��Ӧ�����ڽ��д˲����������ݿ���־����<BR>
1���������ݿ�ģʽΪ��ģʽ����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼-->ѡ��������ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�����Ҽ�ѡ������-->ѡ��ѡ��-->�ڹ��ϻ�ԭ��ģʽ��ѡ�񡰼򵥡���Ȼ��ȷ������<BR>
2���ڵ�ǰ���ݿ��ϵ��Ҽ��������������е��������ݿ⣬һ�������Ĭ�����ò��õ�����ֱ�ӵ�ȷ��<BR>
3��<font color=blue>�������ݿ���ɺ󣬽��齫�������ݿ�������������Ϊ��׼ģʽ����������ͬ��һ�㣬��Ϊ��־��һЩ�쳣����������ǻָ����ݿ����Ҫ����</font>
<BR><BR>

<B>�ġ��趨ÿ���Զ��������ݿ�</B><BR><BR>
<font color=red>ǿ�ҽ������������û����д˲�����</font><BR>
1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����<BR>
2��Ȼ�������˵��еĹ���-->ѡ�����ݿ�ά���ƻ���<BR>
3����һ��ѡ��Ҫ�����Զ����ݵ�����-->��һ�����������Ż���Ϣ������һ�㲻����ѡ��-->��һ��������������ԣ�Ҳһ�㲻ѡ��<BR>
4����һ��ָ�����ݿ�ά���ƻ���Ĭ�ϵ���1�ܱ���һ�Σ��������ѡ��ÿ�챸�ݺ��ȷ��<BR>
5����һ��ָ�����ݵĴ���Ŀ¼��ѡ��ָ��Ŀ¼������������D���½�һ��Ŀ¼�磺d:\databak��Ȼ��������ѡ��ʹ�ô�Ŀ¼������������ݿ�Ƚ϶����ѡ��Ϊÿ�����ݿ⽨����Ŀ¼��Ȼ��ѡ��ɾ�����ڶ�����ǰ�ı��ݣ�һ���趨4��7�죬�⿴���ľ��屸��Ҫ�󣬱����ļ���չ��һ�㶼��bak����Ĭ�ϵ�<BR>
6����һ��ָ��������־���ݼƻ�����������Ҫ��ѡ��-->��һ��Ҫ���ɵı�����һ�㲻��ѡ��-->��һ��ά���ƻ���ʷ��¼�������Ĭ�ϵ�ѡ��-->��һ�����<BR>
7����ɺ�ϵͳ�ܿ��ܻ���ʾSql Server Agent����δ�������ȵ�ȷ����ɼƻ��趨��Ȼ���ҵ��������ұ�״̬���е�SQL��ɫͼ�꣬˫���㿪���ڷ�����ѡ��Sql Server Agent��Ȼ�������м�ͷ��ѡ���·��ĵ�����OSʱ�Զ���������<BR>
8�����ʱ�����ݿ�ƻ��Ѿ��ɹ��������ˣ�������������������ý����Զ�����
<BR><BR>
�޸ļƻ���<BR>
1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->����-->���ݿ�ά���ƻ�-->�򿪺�ɿ������趨�ļƻ������Խ����޸Ļ���ɾ������
<BR><BR>
<B>�塢���ݵ�ת�ƣ��½����ݿ��ת�Ʒ�������</B><BR><BR>
һ������£����ʹ�ñ��ݺͻ�ԭ����������ת�����ݣ�����������£������õ��뵼���ķ�ʽ����ת�ƣ�������ܵľ��ǵ��뵼����ʽ�����뵼����ʽת������һ�����þ��ǿ������������ݿ���Ч�������������С�����������ݿ�Ĵ�С��������Ĭ��Ϊ����SQL�Ĳ�����һ�����˽⣬��������еĲ��ֲ��������⣬������ѯ���������Ա���߲�ѯ��������<BR>
1����ԭ���ݿ�����б����洢���̵�����һ��SQL�ļ���������ʱ��ע����ѡ����ѡ���д�����ű��ͱ�д�����������Ĭ��ֵ�ͼ��Լ���ű�ѡ��<BR>
2���½����ݿ⣬���½����ݿ�ִ�е�һ������������SQL�ļ�<BR>
3����SQL�ĵ��뵼����ʽ���������ݿ⵼��ԭ���ݿ��е����б�����<BR>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
end sub

'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = True
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
Function MakeNewsDir(foldername)
	dim f
	 MakeNewsDir = False
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function

'=====================ϵͳ�ռ����=========================
Sub ShowSpaceInfo(drvpath)
	dim fso,d,size,showsize
	set fso=Dvbbs.iCreateObject("scripting.filesystemobject") 		
	drvpath=server.mappath(drvpath) 		 		
	set d=fso.getfolder(drvpath) 		
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
End Sub	
 	
Sub Showspecialspaceinfo(method)
	dim fso,d,fc,f1,size,showsize,drvpath 		
	set fso=Dvbbs.iCreateObject("scripting.filesystemobject")
	drvpath=server.mappath("../index.asp")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath)
	if method="All" then 		
		size=d.size
	elseif method="Program" then
		set fc=d.Files
		for each f1 in fc
			size=size+f1.size
		next	
	end if
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
end sub 	 	 	
	
Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=Dvbbs.iCreateObject("scripting.filesystemobject")
	drvpathroot=server.mappath("../index.asp")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	drvpath=server.mappath(drvpath)
	if fso.FolderExists(drvpath) then		
		set d=fso.getfolder(drvpath)
		size=d.size
	End If
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 	
 	
Function Drawspecialbar()
	dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
	set fso=Dvbbs.iCreateObject("scripting.filesystemobject")
	drvpathroot=server.mappath("../index.asp")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	set fc=d.files
	for each f1 in fc
		size=size+f1.size
	next
	barsize=cint((size/totalsize)*400)
	Drawspecialbar=barsize
End Function
	
Function GetAllSpace()
	Dim fso,drvpath,d,size
	set fso=Dvbbs.iCreateObject("scripting.filesystemobject")
	drvpath=server.mappath("../index.asp")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath)	
	size=d.size
	set fso=nothing
	GetAllSpace = size
End Function

Function GetFileSize(FileName)
	Dim fso,drvpath,d,size,showsize
	set fso=Dvbbs.iCreateObject("scripting.filesystemobject")
	drvpath=server.mappath(FileName)
	set d=fso.getfile(drvpath)	
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	set fso=nothing
	GetFileSize = showsize
End Function

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

Sub GetSysInfo()
	On Error Resume Next
	Dim WshShell,WshSysEnv
	Set WshShell = Dvbbs.iCreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("SYSTEM")
	okOS = Cstr(WshSysEnv("OS"))
	okCPUS = Cstr(WshSysEnv("NUMBER_OF_PROCESSORS"))
	okCPU = Cstr(WshSysEnv("PROCESSOR_IDENTIFIER"))
	If IsNull(okCPUS) Then
		okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
	ElseIf okCPUS="" Then
		okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
	End If
	If Request.ServerVariables("OS")="" Then okOS=okOS & "(������ Windows Server 2003)"
End Sub
%>