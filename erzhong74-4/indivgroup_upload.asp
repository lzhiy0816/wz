<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!-- #include File="inc/Upload_Class.asp" -->
<%
Const UploadFileType = "gif|jpg|jpeg|bmp|png|rar|txt|zip|mid"

Dim groupboardid
groupboardid = Cint(Request("groupboardid"))

If Request("t")="1" Then
	Upfile_Main()
Else
	Main()
End If
Dvbbs.PageEnd()
Sub Main()
	Dvbbs.Loadtemplates("")
	Dvbbs.Head()
	Dvbbs.ShowErr
	Dim dateupnum
	Dim upset
	upset=CInt(Dvbbs.Groupsetting(66))
	dateupnum=Cint(Dvbbs.UserToday(2))
%>
<script>
if (top.location==self.location){
	top.location="index.asp"
}
</script>
<table border=0 cellspacing=0 cellpadding=0 style="width:100%;height:100%">
<form name="form" method="post" action="IndivGroup_Upload.asp?t=1&groupboardid=<%=groupboardid%>" enctype="multipart/form-data">
<tr>
<%
If Cint(Dvbbs.Groupsetting(7))=0 Then
	Response.Write "<TD class=tablebody2 valign=top>��û���ڱ���̳�ϴ��ļ���Ȩ��!</td>"
Else
	Dim PostRanNum
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("UploadCode") = Cstr(PostRanNum)
%>
<INPUT TYPE="hidden" NAME="UploadCode" value="<%=PostRanNum%>">
<TD id="upid" class=tablebody1 valign=top>
	<input type="file" name="file1" width=200 value="" size="40" /></TD>
<td class=tablebody1 valign=top width=1>
	<input type="submit" name="Submit" value="�ϴ�" onclick="parent.document.PostForm.submit.disabled=true;">
</td>
<td class=tablebody1 valign=top>
	<font color=<%=Dvbbs.mainsetting(1) %> >���컹���ϴ�<%=Dvbbs.Groupsetting(50)-dateupnum%>��</font>��
	<a style="CURSOR: help" title="��̳����:һ��<%=Dvbbs.Groupsetting(40)%>����һ��<%=Dvbbs.Groupsetting(50)%>��,ÿ��<%=Dvbbs.Groupsetting(44)%>K">(�鿴��̳����)</a>
</TD>
<%end if%>
</tr>
</form>
</table>
</body>
</html>
<%
End Sub

Sub Upfile_Main()
	Dvbbs.Loadtemplates("")
	Dvbbs.Head()
	Dvbbs.ShowErr()
	Server.ScriptTimeOut=999999'Ҫ�������̳֧���ϴ����ļ��Ƚϴ󣬾ͱ������á�
	'-----------------------------------------------------------------------------
	'�ύ��֤
	'-----------------------------------------------------------------------------
	If Not Dvbbs.ChkPost Then
		Response.End
	End If
	If Dvbbs.Userid=0 Then
		Response.write "�㻹δ��½��"
		Response.End
	End If
	If Cint(Dvbbs.GroupSetting(7))=0 then
		Response.write "��û���ڱ���̳�ϴ��ļ���Ȩ��"
		Response.End
	End If
%>
	<script>
	parent.document.PostForm.submit.disabled=false;
	</script>
	<table width="100%" height="100%" border=0 cellspacing=0 cellpadding=0>
	<tr><td class=tablebody2 valign=top height=40>
	<%
	UploadFile
	%>
	</td></tr>
	</table>
</body>
</html>
<%
End Sub

Sub UploadFile()
	Dim Upload,FormName,FilePath,ChildFilePath
	Dim File,F_FileName,F_ViewName,F_Filesize,F_FileExt,F_Type
	Dim Previewpath,DrawInfo,InceptMaxFile
	Dim OnceUPCount
	OnceUPCount = Request.Cookies("upNum")
	If OnceUPCount = "" or Not Isnumeric(OnceUPCount) Then
		OnceUPCount = 0
	Else
		OnceUPCount = Clng(OnceUPCount)
	End If
	If OnceUPCount >= Clng(Dvbbs.GroupSetting(40)) then
 		Response.write "һ��ֻ���ϴ�"&Dvbbs.GroupSetting(40)&"���ļ���"
		Exit Sub
	Else
		InceptMaxFile = Clng(Dvbbs.GroupSetting(40)) - OnceUPCount
	End If
	If Not IsNumeric(Dvbbs.UserToday(2)) Then Dvbbs.UserToday(2) = 0
	If Clng(Dvbbs.UserToday(2))>Clng(Dvbbs.GroupSetting(50)) Then
 		Response.write "�ѳ�����������̳ÿ���ϴ����ļ�����"&Dvbbs.GroupSetting(50)&"����"
		Exit Sub
	Else
		If Clng(Dvbbs.GroupSetting(50))-Clng(Dvbbs.UserToday(2))<InceptMaxFile Then
			InceptMaxFile = Clng(Dvbbs.GroupSetting(50))-Clng(Dvbbs.UserToday(2))
		End If
	End If

	'�ϴ�Ŀ¼
	FilePath = CreatePath(CheckFolder)
	'����ϵͳ�ϴ�Ŀ¼���¼�Ŀ¼·��
	ChildFilePath = Replace(FilePath,CheckFolder,"")
	'Ԥ��ͼƬĿ¼·��
	Previewpath = "PreviewImage/"
	Previewpath = CreatePath(Previewpath)

	If Dvbbs.Forum_UploadSetting(17)="1" Then
		DrawInfo = Dvbbs.Forum_UploadSetting(4)
	ElseIf Dvbbs.Forum_UploadSetting(17)="2" Then
		DrawInfo = Dvbbs.Forum_UploadSetting(9)
	Else
		DrawInfo = ""
	End If
	If DrawInfo = "0" Then
		DrawInfo = ""
		Dvbbs.Forum_UploadSetting(17) = 0
	End If

	Set Upload = New UpFile_Cls
	Upload.UploadType			= Cint(Dvbbs.Forum_UploadSetting(2))	'�����ϴ��������
	Upload.UploadPath			= FilePath								'�����ϴ�·��
	Upload.InceptFileType		= Replace(UploadFileType,"|",",")		'�����ϴ��ļ�����
	Upload.MaxSize				= Int(Dvbbs.GroupSetting(44))			'��λ KB
	Upload.InceptMaxFile		= InceptMaxFile							'ÿ���ϴ��ļ���������
	Upload.ChkSessionName		= "UploadCode"							'��ֹ�ظ��ύ��SESSION�����ύ�ı�Ҫһ�¡�
	'Ԥ��ͼƬ����
	Upload.PreviewType			= Cint(Dvbbs.Forum_UploadSetting(3))	'����Ԥ��ͼƬ�������
	Upload.PreviewImageWidth	= Dvbbs.Forum_UploadSetting(14)			'����Ԥ��ͼƬ���
	Upload.PreviewImageHeight	= Dvbbs.Forum_UploadSetting(15)			'����Ԥ��ͼƬ�߶�
	Upload.DrawImageWidth		= Dvbbs.Forum_UploadSetting(11)			'����ˮӡͼƬ������������
	Upload.DrawImageHeight		= Dvbbs.Forum_UploadSetting(12)			'����ˮӡͼƬ����������߶�
	Upload.DrawGraph			= Dvbbs.Forum_UploadSetting(10)			'����ˮӡ͸����
	Upload.DrawFontColor		= Dvbbs.Forum_UploadSetting(6)			'����ˮӡ������ɫ
	Upload.DrawFontFamily		= Dvbbs.Forum_UploadSetting(7)			'����ˮӡ���������ʽ
	Upload.DrawFontSize			= Dvbbs.Forum_UploadSetting(5)			'����ˮӡ���������С
	Upload.DrawFontBold			= Dvbbs.Forum_UploadSetting(8)			'����ˮӡ�����Ƿ����
	Upload.DrawInfo				= DrawInfo								'����ˮӡ������Ϣ��ͼƬ��Ϣ
	Upload.DrawType				= Dvbbs.Forum_UploadSetting(17)			'0=������ˮӡ ��1=����ˮӡ���֣�2=����ˮӡͼƬ
	Upload.DrawXYType			= Dvbbs.Forum_UploadSetting(13)			'"0" =���ϣ�"1"=����,"2"=����,"3"=����,"4"=����
	Upload.DrawSizeType			= Dvbbs.Forum_UploadSetting(16)			'"0"=�̶���С��"1"=�ȱ�����С
	If Dvbbs.Forum_UploadSetting(18)<>"" or Dvbbs.Forum_UploadSetting(18)<>"0" Then
		Upload.TransitionColor	= Dvbbs.Forum_UploadSetting(18)			'͸������ɫ����
	End If
	'ִ���ϴ�
	Upload.SaveUpFile
	If Upload.ErrCodes<>0 Then
		Response.write "����"& Upload.Description & "[ <a href=""IndivGroup_Upload.asp?groupboardid=" & groupboardid & """>�����ϴ�</a> ]"
		Exit Sub
	End If
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
				F_FileName = FilePath & File.FileName
				'����Ԥ����ˮӡͼƬ
				If Upload.PreviewType<>999 and File.FileType=1 then
						F_Viewname = Previewpath & "pre" & Replace(File.FileName,File.FileExt,"") & "jpg"
						'����Ԥ��ͼƬ:Call CreateView(ԭʼ�ļ���·��,Ԥ���ļ�����·��,ԭ�ļ���׺)
						Upload.CreateView F_FileName,F_Viewname,File.FileExt
				End If
				UploadSave F_FileName,ChildFilePath&File.FileName,File.FileExt,F_Viewname,File.FileSize,File.FileType,File.FileOldName
			Set File = Nothing
		Next
	Else
		Response.write "����ȷѡ��Ҫ�ϴ����ļ���[ <a href=""IndivGroup_Upload.asp?groupboardid=" & groupboardid & """>�����ϴ�</a> ]"
		Exit Sub
	End If
	Call Suc_upload(Upload.Count,OnceUPCount)
	Set Upload = Nothing
End Sub

'�����ϴ����ݲ����ظ���ID
Sub UploadSave(FileName,ChildFileName,FileExt,ViewName,FileSize,F_Type,F_OldName)
	Dim ShwoFileName
	ShwoFileName = Dvbbs.Checkstr(Replace(FileName,CheckFolder,"UploadFile/"))
	ChildFileName = Dvbbs.Checkstr(ChildFileName)
	Dvbbs.Execute("Insert into Dv_upFile (F_BoardID,F_UserID,F_Username,F_Filename,F_Viewname,F_FileType,F_Type,F_FileSize,F_Readme,F_Flag,F_OldName) values ("&groupboardid&","&Dvbbs.UserID&",'"&Dvbbs.Membername&"','"&ChildFileName&"','"&ViewName&"','"&FileExt&"',"&F_Type&","&FileSize&",'Ȧ���ϴ�ͼƬ',4,'"&Left(Dvbbs.Checkstr(F_OldName),50)&"')")
	Dim Rs,UpFileID,DownloadID
	Set Rs=Dvbbs.Execute("Select top 1 F_ID From Dv_upFile order by F_ID desc")
		DownloadID=rs(0)
		UpFileID = DownloadID & ","
	Rs.Close
	Set Rs=nothing
	If F_Type=1 or F_Type=2 then
		Response.write "<script>var str = '<br />[upload="&FileExt&","&F_OldName&"]"&ShwoFileName&"[/upload]<br>'</script>"
	Else
		Response.write "<script>var str ='<br />[upload="&FileExt&","&F_OldName&"]viewFile.asp?ID="&DownloadID&"[/upload]<br>'</script>"
	End If
	Response.write "<script>parent.dvtextarea.insert(str);</script>"
End Sub

Sub Suc_upload(UpCount,upNum)
	upNum = upNum + UpCount
	Response.Cookies("upNum") = upNum
	Dim iUserInfo
	Dvbbs.UserToday(2) = Dvbbs.UserToday(2)+UpCount
	iUserInfo = Dvbbs.UserToday(0) & "|" & Dvbbs.UserToday(1) & "|" & Dvbbs.UserToday(2) & "|" & Dvbbs.UserToday(3) & "|" & Dvbbs.UserToday(4)
	iUserInfo=Dvbbs.Checkstr(iUserInfo)
	If upNum < Clng(Dvbbs.GroupSetting(40)) And Dvbbs.UserToday(2) < Clng(Dvbbs.GroupSetting(50)) Then
		Response.Write UpCount & "���ļ��ϴ��ɹ�,Ŀǰ�����ܹ��ϴ���" & Dvbbs.UserToday(2) & "������ [ <a href=IndivGroup_Upload.asp?groupboardid=" & groupboardid & ">�����ϴ�</a> ]"
	Else
		Response.write UpCount & "���ļ��ϴ��ɹ�!�����Ѵﵽ�ϴ������ޡ�"
	End If
	Dvbbs.Execute("UPDATE [Dv_user] SET UserToday = '" & iUserInfo &"' WHERE UserID = " & Dvbbs.UserID)
	Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usertoday").text=iUserInfo
End Sub


'��ȡ�ϴ�Ŀ¼
Function CheckFolder()
	If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
	CheckFolder = Replace(Replace(Dvbbs.Forum_Setting(76),Chr(0),""),".","")
	'��Ŀ¼���(/)
	If Right(CheckFolder,1)<>"/" Then CheckFolder=CheckFolder&"/"
End Function

'���·��Զ������ϴ��ļ���,��Ҫ�ƣӣ����֧�֡�
Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	'�����´����ϴ��ļ��У���ʽ��2003��8
	uploadpath = year(now) & "-" & month(now)
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	On Error Resume Next
	Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & uploadpath))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & uploadpath)
		End If
		If Err.Number = 0 Then
			CreatePath = PathValue & uploadpath & "/"
		Else
			CreatePath = PathValue
		End If
	Set objFSO = Nothing
End Function
%>